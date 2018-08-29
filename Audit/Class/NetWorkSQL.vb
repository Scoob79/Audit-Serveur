Imports System.IO
Imports System.Data.OleDb
Imports System.Data.SqlClient
Public Class SQL
    Public Const ChaineDeConnexion As String = "Provider=microsoft.jet.oledb.4.0;Data Source=D:\Users\u165147\source\repos\Audit\Audit\BDD\BDD.mdb"
    Public Function Requete(ChaineDeConnexion As String, Chaine As String) As String
        Dim connect As New OleDbConnection(ChaineDeConnexion)
        Dim cmd As New OleDbCommand
        Try
            connect.Open()
            cmd.Connection = connect
            cmd.CommandText = Replace(Replace(Replace(Chaine, "'", " "), "�", ""), ".", " ")
            cmd.ExecuteNonQuery()
            connect.Close()
        Catch ex As Exception
            Requete = ex.Message
            MsgBox(ex.Message)
        End Try

    End Function

    Public Sub ChargementFichier(FichierAudit As String)
        Dim Fichier As New StreamReader(FichierAudit)
        Dim Ligne As String, Lecteur As String, NS As String, Tipe As String, SysFic As String, EspLibre As String, EspTotal As String, Pourcentage As Integer
        Dim Nom As String, Description As String, Statut As String, Etat As String, CodeSortie As String, X As Integer, Data As String, ColData(12) As String
        Dim NomPoste As String, DescPoste As String, OS As String, Version As String, DateInstall As String, NumDernierSPMa As String, NumDernierSPMi As String, Fabricant As String
        Dim Model As String, Manufacturier As String, Modèle As String, TypeProc As String, NomProc As String, DescProc As String, VitesseACT As String
        Dim VitesseMAX As String, Taille As String, N As String, NomCarte As String, TypeCarte As String, MAC As String, RxVitesseMAX As String, IP(20) As String
        Dim MSR As String, DHCP As String, AddDHCP As String, DNS As String, Utilisateurs As String, Groupes As String, Expiration As String, MDPVieMin As String
        Dim MDPVieMax As String, MDPLongueur As String, MDPAnterieur As String, SeuilVerrou As String, DureeVerrou As String, FenObsVerrou As String, RolePoste As String, Logiciels As String
        Dim Pilotes As String, service(5) As String, MAJ(6) As String, Retour As String, Ret As String, ADDIP As String, a As Integer

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
                If InStr(Ligne, "DateInstall") > 0 Then DateInstall = Ligne.Substring(InStr(Ligne, "="), 14)
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
        DateInstall = DateInstall.Substring(6, 2) & "/" & DateInstall.Substring(4, 2) & "/" & DateInstall.Substring(0, 4)
#Disable Warning BC42104 ' La variable est utilisée avant de se voir attribuer une valeur
        Retour = Requete(ChaineDeConnexion, "INSERT INTO Serveur (POSTE_NomPoste, POSTE_DescPoste, POSTE_OS, POSTE_Version, POSTE_DateInstall, POSTE_NumDernierSPMa, POSTE_NumDernierSPMi ," &
                "POSTE_Fabricant, POSTE_Model) values ('" & NomPoste & "', '" & DescPoste & "', '" & OS & "', '" & Version & "', '" & DateInstall & "', '" & NumDernierSPMa _
& "', '" & NumDernierSPMi & "', '" & Fabricant & "', '" & Model & "')")
        If Retour <> "" Then
            Retour = Requete(ChaineDeConnexion, "UPDATE Serveur SET POSTE_NomPoste='" & NomPoste & "', POSTE_DescPoste='" & DescPoste & "', POSTE_OS='" & OS &
            "', POSTE_Version='" & Version & "', POSTE_DateInstall='" & DateInstall & "', POSTE_NumDernierSPMa='" & NumDernierSPMa & "', POSTE_NumDernierSPMi='" & NumDernierSPMi & "', POSTE_Fabricant='" & Fabricant & "', POSTE_Model='" & Model & "'")
        End If
#Enable Warning BC42104 ' La variable est utilisée avant de se voir attribuer une valeur

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
        Retour = Requete(ChaineDeConnexion, "UPDATE Serveur SET CARTE_MERE_Nom='" & Nom & "', CARTE_MERE_Modèle='" & Modèle & "', CARTE_MERE_Manufacturier='" & Manufacturier & "' WHERE  (((Serveur.POSTE_NomPoste)='" & NomPoste & "'));")

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
        Retour = Requete(ChaineDeConnexion, "UPDATE Serveur SET PROCESSEUR_TypeProc='" & Nom & "', PROCESSEUR_NomProc='" & Modèle & "', PROCESSEUR_DescProc='" & Manufacturier _
            & "', PROCESSEUR_VitesseACT='" & VitesseACT & "', PROCESSEUR_VitesseMAX='" & VitesseMAX & "' WHERE  (((Serveur.POSTE_NomPoste)='" & NomPoste & "'));")

        Do
            Ligne = Fichier.ReadLine
            If Ligne = "[MEMOIRE]" Then
                Ligne = Fichier.ReadLine
                Ligne = Fichier.ReadLine
                If InStr(Ligne, "Taille") > 0 Then Taille = Format(Val(Ligne.Substring(InStr(Ligne, "="))) / 1024, "# ### ###.00 Mo")
                Exit Do
            End If
        Loop Until Ligne Is Nothing
        Retour = Requete(ChaineDeConnexion, "UPDATE Serveur SET MEMOIRE_Taille='" & Taille & "' WHERE  (((Serveur.POSTE_NomPoste)='" & NomPoste & "'));")

        N = -1
        Do
            Ligne = Fichier.ReadLine
            If Ligne = "[HDD]" Then
                Retour = Requete(ChaineDeConnexion, "CREATE TABLE " & Replace(NomPoste, "-", "_") & "_HDD (HDD_Lecteur VARCHAR(20) PRIMARY KEY, HDD_NS VARCHAR(20), HDD_Type VARCHAR(40), HDD_SysFic VARCHAR(20), HDD_EspLibre VARCHAR(20), HDD_EspTotal VARCHAR(20))")
                Ligne = Fichier.ReadLine
                Do While Not Ligne = "[RESEAU]"
                    Ligne = Fichier.ReadLine
                    If InStr(Ligne, "Lecteur") > 0 Then Lecteur = Ligne.Substring(InStr(Ligne, "="))
                    If Ligne = "[RESEAU]" Then Exit Do

                    Ligne = Fichier.ReadLine
                    If InStr(Ligne, "NS") > 0 Then NS = Ligne.Substring(InStr(Ligne, "="))
                    If Ligne = "[RESEAU]" Then Exit Do

                    Ligne = Fichier.ReadLine
                    If InStr(Ligne, "Type") > 0 Then Tipe = Ligne.Substring(InStr(Ligne, "="))
                    If Ligne = "[RESEAU]" Then Exit Do

                    Ligne = Fichier.ReadLine
                    If InStr(Ligne, "SysFic") > 0 Then SysFic = Ligne.Substring(InStr(Ligne, "="))
                    If Ligne = "[RESEAU]" Then Exit Do

                    Ligne = Fichier.ReadLine
                    If InStr(Ligne, "EspLibre") > 0 Then EspLibre = LTrim(Ligne.Substring(InStr(Ligne, "=")))
                    If Ligne = "[RESEAU]" Then Exit Do

                    Ligne = Fichier.ReadLine
                    If InStr(Ligne, "EspTotal") > 0 Then EspTotal = LTrim(Ligne.Substring(InStr(Ligne, "=")))
                    If Ligne = "[RESEAU]" Then Exit Do

                    If Val(EspLibre) > 1024 Then
                        EspLibre = LTrim(Format(Val(EspLibre) / 1024, "# ### ###.00 Go"))
                    Else
                        EspLibre = EspLibre.Substring(InStr(EspLibre, "=")) & " Mo"
                    End If
                    If Val(EspTotal) > 1024 Then
                        EspTotal = LTrim(Format(Val(EspTotal) / 1024, "# ### ###.00 Go"))
                    Else
                        EspTotal = EspTotal & " Mo"
                    End If
                    Pourcentage = Val(EspLibre) / (Val(EspTotal) / 100)
                    N += 1
                    Ligne = Fichier.ReadLine
                    If Retour <> "" Then
                        Ret = Requete(ChaineDeConnexion, "UPDATE " & Replace(NomPoste, "-", "_") & "_HDD SET HDD_Lecteur='" &
                            Lecteur & "', HDD_NS='" & NS & "', HDD_Type='" & Tipe & "', HDD_SysFic='" & SysFic & "', HDD_EspLibre='" & EspLibre & "', HDD_EspTotal='" & EspTotal & "' WHERE HDD_Lecteur='" & Lecteur & "';")
                    Else
                        Ret = Requete(ChaineDeConnexion, "INSERT INTO " & Replace(NomPoste, "-", "_") & "_HDD (HDD_Lecteur, HDD_NS, HDD_Type, HDD_SysFic, HDD_EspLibre, HDD_EspTotal) values ('" &
                            Lecteur & "', '" & NS & "', '" & Tipe & "', '" & SysFic & "', '" & EspLibre & "', '" & EspTotal & "')")
                    End If
                Loop
                Exit Do
            End If
        Loop Until Ligne Is Nothing
        N = -1
        MAC = "temp"
        Retour = Requete(ChaineDeConnexion, "CREATE TABLE " & Replace(NomPoste, "-", "_") & "_Reseau (RESEAU_NomCarte VARCHAR(255), RESEAU_TypeCarte VARCHAR(20), RESEAU_Description" &
                             " LONGTEXT, RESEAU_MAC VARCHAR(20) PRIMARY KEY, RESEAU_VitesseMAX VARCHAR(20), RESEAU_IP VARCHAR(20), RESEAU_MSR VARCHAR(20), RESEAU_DHCP VARCHAR(20)" &
                             ", RESEAU_ADDDHCP VARCHAR(20), RESEAU_DNS VARCHAR(20))")
        If Retour <> "" Then
            Requete(ChaineDeConnexion, "DROP TABLE " & Replace(NomPoste, "-", "_") & "_Reseau")
            Retour = Requete(ChaineDeConnexion, "CREATE TABLE " & Replace(NomPoste, "-", "_") & "_Reseau (RESEAU_NomCarte VARCHAR(255), RESEAU_TypeCarte VARCHAR(20), RESEAU_Description" &
                                 " LONGTEXT, RESEAU_MAC VARCHAR(20) PRIMARY KEY, RESEAU_VitesseMAX VARCHAR(20), RESEAU_IP VARCHAR(20), RESEAU_MSR VARCHAR(20), RESEAU_DHCP VARCHAR(20)" &
                                 ", RESEAU_ADDDHCP VARCHAR(20), RESEAU_DNS VARCHAR(20))")
        End If
        Do
            Ligne = Fichier.ReadLine
            If InStr(Ligne, "NomCarte") <> 0 Then MAC = "temp" : Requete(ChaineDeConnexion, "INSERT INTO " & Replace(NomPoste, "-", "_") & "_RESEAU (RESEAU_NomCarte, RESEAU_MAC) VALUES ('" & Ligne.Substring(InStr(Ligne, "=")) & "','" & MAC & "')")
            If InStr(Ligne, "TypeCarte") <> 0 Then Requete(ChaineDeConnexion, "UPDATE " & Replace(NomPoste, "-", "_") & "_RESEAU SET RESEAU_TypeCarte='" & Ligne.Substring(InStr(Ligne, "=")) & "' WHERE RESEAU_MAC='" & MAC & "'")
            If InStr(Ligne, "Description") <> 0 Then Requete(ChaineDeConnexion, "UPDATE " & Replace(NomPoste, "-", "_") & "_RESEAU SET RESEAU_Description='" & Ligne.Substring(InStr(Ligne, "=")) & "' WHERE RESEAU_MAC='" & MAC & "'")
            If InStr(Ligne, "@MAC") <> 0 Then
                Requete(ChaineDeConnexion, "UPDATE " & Replace(NomPoste, "-", "_") & "_RESEAU SET RESEAU_MAC='" & Ligne.Substring(InStr(Ligne, "=")) & "' WHERE RESEAU_MAC='" & MAC & "'")
                MAC = Ligne.Substring(InStr(Ligne, "="))
            End If
            If InStr(Ligne, "VitesseMAX") <> 0 Then Requete(ChaineDeConnexion, "UPDATE " & Replace(NomPoste, "-", "_") & "_RESEAU SET RESEAU_VitesseMAX='" & Ligne.Substring(InStr(Ligne, "=")) & "' WHERE RESEAU_MAC='" & MAC & "'")
            If InStr(Ligne, "@IP") <> 0 Then
                If Ligne.Substring(InStr(Ligne, "=")) = "" Then
                    N = -1
                Else
                    N += 1
                    ADDIP = Ligne.Substring(InStr(Ligne, "=")) & "|" & ADDIP
                    Requete(ChaineDeConnexion, "UPDATE " & Replace(NomPoste, "-", "_") & "_RESEAU SET RESEAU_IP='" & ADDIP & "' WHERE RESEAU_MAC='temp'")
                End If
            End If
            If InStr(Ligne, "MSR") <> 0 Then Requete(ChaineDeConnexion, "UPDATE " & Replace(NomPoste, "-", "_") & "_RESEAU SET RESEAU_MSR='" & Ligne.Substring(InStr(Ligne, "=")) & "' WHERE RESEAU_MAC='" & MAC & "'")
            If InStr(Ligne, "DHCP") <> 0 And InStr(Ligne, "@") = 0 Then Requete(ChaineDeConnexion, "UPDATE " & Replace(NomPoste, "-", "_") & "_RESEAU SET RESEAU_DHCP='" & Ligne.Substring(InStr(Ligne, "=")) & "' WHERE RESEAU_MAC='" & MAC & "'")
            If InStr(Ligne, "@DHCP") <> 0 Then Requete(ChaineDeConnexion, "UPDATE " & Replace(NomPoste, "-", "_") & "_RESEAU SET RESEAU_ADDDHCP='" & Ligne.Substring(InStr(Ligne, "=")) & "' WHERE RESEAU_MAC='" & MAC & "'")
            If InStr(Ligne, "@DNS") <> 0 Then Requete(ChaineDeConnexion, "UPDATE " & Replace(NomPoste, "-", "_") & "_RESEAU SET RESEAU_DNS='" & Ligne.Substring(InStr(Ligne, "=")) & "' WHERE RESEAU_MAC='" & MAC & "'")
            If Ligne = "[UTILISATEURS]" Then Exit Do
        Loop

        N = 0
        Do
            Ligne = Fichier.ReadLine
            If Ligne <> "" And Ligne <> "[GROUPES]" Then Utilisateurs = Utilisateurs & "|" & Ligne
        Loop While Not Ligne = "[GROUPES]"
        Requete(ChaineDeConnexion, "UPDATE Serveur SET UTILISATEURS='" & Utilisateurs & "' WHERE POSTE_NomPoste='" & NomPoste & "'")

        N = 0
        Do
            Ligne = Fichier.ReadLine
            If Ligne <> "" And Ligne <> "[STRATEGIE]" Then Groupes = Groupes & "|" & Ligne
            Groupes = Replace(Replace(Groupes, "�", ""), "'", " ")
        Loop While Not Ligne = "[STRATEGIE]"
        Requete(ChaineDeConnexion, "UPDATE Serveur SET GROUPES='" & Groupes & "' WHERE POSTE_NomPoste='" & NomPoste & "'")

        Do
            If Ligne = "[STRATEGIE]" Then
                Ligne = Fichier.ReadLine
                Ligne = Fichier.ReadLine
                If InStr(Ligne, "Expiration") > 0 Then Requete(ChaineDeConnexion, "UPDATE Serveur SET STRATEGIE_Expiration='" & Ligne.Substring(InStr(Ligne, "=")) & "' WHERE POSTE_NomPoste='" & NomPoste & "'")
                Ligne = Fichier.ReadLine
                If InStr(Ligne, "MDPVieMin") > 0 Then Requete(ChaineDeConnexion, "UPDATE Serveur SET STRATEGIE_MDPVieMin='" & Ligne.Substring(InStr(Ligne, "=")) & "' WHERE POSTE_NomPoste='" & NomPoste & "'")
                Ligne = Fichier.ReadLine
                If InStr(Ligne, "MDPVieMax") > 0 Then Requete(ChaineDeConnexion, "UPDATE Serveur SET STRATEGIE_MDPVieMax='" & Ligne.Substring(InStr(Ligne, "=")) & "' WHERE POSTE_NomPoste='" & NomPoste & "'")
                Ligne = Fichier.ReadLine
                If InStr(Ligne, "MDPLongueur") > 0 Then Requete(ChaineDeConnexion, "UPDATE Serveur SET STRATEGIE_MDPLongueur='" & Ligne.Substring(InStr(Ligne, "=")) & "' WHERE POSTE_NomPoste='" & NomPoste & "'")
                Ligne = Fichier.ReadLine
                If InStr(Ligne, "MDPAnterieur") > 0 Then Requete(ChaineDeConnexion, "UPDATE Serveur SET STRATEGIE_MDPAnterieur='" & Ligne.Substring(InStr(Ligne, "=")) & "' WHERE POSTE_NomPoste='" & NomPoste & "'")
                Ligne = Fichier.ReadLine
                If InStr(Ligne, "SeuilVerrou") > 0 Then Requete(ChaineDeConnexion, "UPDATE Serveur SET STRATEGIE_SeuilVerrou='" & Ligne.Substring(InStr(Ligne, "=")) & "' WHERE POSTE_NomPoste='" & NomPoste & "'")
                Ligne = Fichier.ReadLine
                If InStr(Ligne, "DureeVerrou") > 0 Then Requete(ChaineDeConnexion, "UPDATE Serveur SET STRATEGIE_DureeVerrou='" & Ligne.Substring(InStr(Ligne, "=")) & "' WHERE POSTE_NomPoste='" & NomPoste & "'")
                Ligne = Fichier.ReadLine
                If InStr(Ligne, "FenObsVerrou") > 0 Then Requete(ChaineDeConnexion, "UPDATE Serveur SET STRATEGIE_FenObsVerrou='" & Ligne.Substring(InStr(Ligne, "=")) & "' WHERE POSTE_NomPoste='" & NomPoste & "'")
                Ligne = Fichier.ReadLine
                If InStr(Ligne, "RolePoste") > 0 Then Requete(ChaineDeConnexion, "UPDATE Serveur SET STRATEGIE_RolePoste='" & Ligne.Substring(InStr(Ligne, "=")) & "' WHERE POSTE_NomPoste='" & NomPoste & "'")
                Ligne = Fichier.ReadLine
                Ligne = Fichier.ReadLine
                Ligne = Fichier.ReadLine
                Exit Do
            End If
        Loop Until Ligne Is Nothing

        N = 0
        Do
            Ligne = Fichier.ReadLine
            If Ligne <> "" And Ligne <> "[PILOTES]" Then N += 1 : Logiciels = Logiciels & "|" & Ligne
            Logiciels = Replace(Replace(Logiciels, "�", ""), "'", " ")
        Loop While Not Ligne = "[PILOTES]"
        Requete(ChaineDeConnexion, "UPDATE Serveur SET Logiciels='" & Logiciels & "' WHERE POSTE_NomPoste='" & NomPoste & "'")

        N = 0
        Do
            Ligne = Fichier.ReadLine
            If Ligne <> "" And Ligne <> "[SERVICES]" Then N += 1 : Pilotes = Pilotes & "|" & Ligne
            Pilotes = Replace(Replace(Pilotes, "�", ""), "'", " ")
        Loop While Not Ligne = "[SERVICES]"
        Requete(ChaineDeConnexion, "UPDATE Serveur SET Pilotes='" & Pilotes & "' WHERE POSTE_NomPoste='" & NomPoste & "'")

        N = 0
        Retour = Requete(ChaineDeConnexion, "CREATE TABLE " & Replace(NomPoste, "-", "_") & "_SERVICES (SERVICES_Nom VARCHAR(255) PRIMARY KEY,SERVICES_Description LONGTEXT,SERVICES_Statut VARCHAR(15),SERVICES_Etat VARCHAR(15),SERVICES_code VARCHAR(10))")
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
                    service(1) = Nom.Substring(4)
                    service(2) = Description.Substring(12)
                    service(3) = Statut.Substring(7)
                    service(4) = Etat.Substring(8)
                    service(5) = CodeSortie.Substring(11)
                    Ligne = Fichier.ReadLine
                    If Retour <> "" Then
                        Ret = Requete(ChaineDeConnexion, "UPDATE " & Replace(NomPoste, "-", "_") & "_SERVICES SET SERVICES_Nom='" &
                            service(1) & "', SERVICES_Description='" & service(2) & "', SERVICES_Statut='" & service(3) & "', SERVICES_Etat='" & service(4) & "', SERVICES_code='" & service(5) & "' WHERE SERVICES_Nom='" & service(1) & "';")
                    Else
                        Ret = Requete(ChaineDeConnexion, "INSERT INTO " & Replace(NomPoste, "-", "_") & "_SERVICES (SERVICES_Nom, SERVICES_Description, SERVICES_Statut, SERVICES_Etat, SERVICES_code) values ('" &
                            service(1) & "', '" & service(2) & "', '" & service(3) & "', '" & service(4) & "', '" & service(5) & "')")
                    End If
                Loop
                Exit Do
            End If
        Loop Until Ligne Is Nothing

        N = 0
        Ligne = Fichier.ReadLine
        Ligne = Fichier.ReadLine
        Retour = Requete(ChaineDeConnexion, "CREATE TABLE " & Replace(NomPoste, "-", "_") & "_MAJ (MAJ_Nom VARCHAR(255) ,MAJ_Poste VARCHAR(255),MAJ_Descrption VARCHAR(15),MAJ_ID VARCHAR(15) PRIMARY KEY,MAJ_InstallPar VARCHAR(10)), MAJ_Date date)")
        Do
            Data = Fichier.ReadLine
            If Data = "" Then Exit Do
            Do
                X = X + 1
                ColData(X) = Data.Substring(0, InStr(Data, "  "))
                Data = LTrim(Data.Substring(Len(ColData(X))))
            Loop While Not Len(Data) = 0
            N += 1
            MAJ(1) = ColData(1)
            MAJ(2) = ColData(2)
            MAJ(3) = ColData(3)
            MAJ(4) = ColData(4)
            MAJ(5) = ColData(5)
            MAJ(6) = ColData(6)
            X = 0
            If Retour <> "" Then
                Ret = Requete(ChaineDeConnexion, "UPDATE " & Replace(NomPoste, "-", "_") & "_MAJ SET MAJ_Nom='" &
                            MAJ(1) & "', MAJ_Poste='" & MAJ(2) & "', MAJ_Descrption='" & MAJ(3) & "', MAJ_ID='" & MAJ(4) & "', MAJ_InstallPar='" & MAJ(5) & "', MAJ_Date='" & MAJ(6) & "' WHERE HDD_Lecteur='" & MAJ(4) & "';")
            Else
                Ret = Requete(ChaineDeConnexion, "INSERT INTO " & Replace(NomPoste, "-", "_") & "_MAJ (MAJ_Nom, MAJ_Poste, MAJ_Descrption, MAJ_ID, MAJ_InstallPar, MAJ_Date) values ('" &
                            MAJ(1) & "', '" & MAJ(2) & "', '" & MAJ(3) & "', '" & MAJ(4) & "', '" & MAJ(5) & "', '" & MAJ(6) & "')")
            End If
        Loop

    End Sub
End Class
