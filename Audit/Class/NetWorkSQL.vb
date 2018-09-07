Imports System.IO
Imports System.Data.OleDb
Imports Microsoft.Office.Interop.Access.Dao

Public Class SQL
    Public Const ChaineDeConnexion As String = "Provider=microsoft.jet.oledb.4.0;Data Source=D:\Users\u165147\source\repos\Audit\Audit\BDD\BDD.mdb"
    Public Function Requete(ChaineDeConnexion As String, Chaine As String) As String
        Dim connect As New OleDbConnection(ChaineDeConnexion)
        Dim cmd As New OleDbCommand
        Try
            connect.Open()
            cmd.Connection = connect
            cmd.CommandText = Chaine
            cmd.ExecuteNonQuery()
            connect.Close()
        Catch ex As Exception
            Requete = ex.Message
            connect.Close()
        End Try

    End Function
    Public Function RequeteQ(ChaineDeConnexion As String, Chaine As String) As OleDbDataReader
        Dim connect As New OleDbConnection(ChaineDeConnexion)
        Dim cmd As New OleDbCommand
        Try
            connect.Open()
            cmd.Connection = connect
            cmd.CommandText = Chaine
            RequeteQ = cmd.ExecuteReader()
            connect.Close()
        Catch ex As Exception
            connect.Close()
        End Try

    End Function

    Public Sub ChargementFichier(FichierAudit As String)
        Dim Fichier As New StreamReader(FichierAudit)
        Dim Ligne As String, Lecteur As String, NS As String, Tipe As String, SysFic As String, EspLibre As String, EspTotal As String, Pourcentage As Integer
        Dim Nom As String, Description As String, Statut As String, Etat As String, CodeSortie As String, X As Integer, Data As String, ColData(12) As String
        Dim NomPoste As String, DescPoste As String, OS As String, Version As String, DateInstall As String, NumDernierSPMa As String, NumDernierSPMi As String, Fabricant As String
        Dim Model As String, Manufacturier As String, Modèle As String, TypeProc As String, NomProc As String, DescProc As String, VitesseACT As String
        Dim VitesseMAX As String, Taille As String, N As String, MAC As String, IP(20) As String, Utilisateurs As String, Groupes As String, Logiciels As String
        Dim Pilotes As String, service(5) As String, MAJ(10) As String, Retour As String, Ret As String, ADDIP As String, CléServeur As String, Tableau(500) As String
        Dim services As String, Res As String, MAJLigne(10) As String

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
        CléServeur = NomPoste & "_" & Format(Now, "ddMMyyyy")
        DateInstall = DateInstall.Substring(6, 2) & "/" & DateInstall.Substring(4, 2) & "/" & DateInstall.Substring(0, 4)
#Disable Warning BC42104 ' La variable est utilisée avant de se voir attribuer une valeur
        Retour = Requete(ChaineDeConnexion, "INSERT INTO Serveur (POSTE_NomPoste, POSTE_DescPoste, POSTE_OS, POSTE_Version, POSTE_DateInstall, POSTE_NumDernierSPMa, POSTE_NumDernierSPMi ," &
                "POSTE_Fabricant, POSTE_Model, DATE_COLLECTE) values ('" & CléServeur & "', '" & DescPoste & "', '" & OS & "', '" & Version & "', '" & DateInstall & "', '" & NumDernierSPMa _
& "', '" & NumDernierSPMi & "', '" & Fabricant & "', '" & Model & "', '" & Now & "')")
        If Retour <> "" Then
            Retour = Requete(ChaineDeConnexion, "UPDATE Serveur SET POSTE_NomPoste='" & CléServeur & "', POSTE_DescPoste='" & DescPoste & "', POSTE_OS='" & OS &
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
                If InStr(Ligne, "Modele") > 0 Then Modèle = Ligne.Substring(InStr(Ligne, "="))
                Ligne = Fichier.ReadLine
                If InStr(Ligne, "Manufacturier") > 0 Then Manufacturier = Ligne.Substring(InStr(Ligne, "="))
                Exit Do
            End If
        Loop Until Ligne Is Nothing
        Retour = Requete(ChaineDeConnexion, "UPDATE Serveur SET CARTE_MERE_Nom='" & Nom & "', CARTE_MERE_Modèle='" & Modèle & "', CARTE_MERE_Manufacturier='" & Manufacturier & "' WHERE  POSTE_NomPoste='" & CléServeur & "';")

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
            & "', PROCESSEUR_VitesseACT='" & VitesseACT & "', PROCESSEUR_VitesseMAX='" & VitesseMAX & "' WHERE  POSTE_NomPoste='" & CléServeur & "';")

        Do
            Ligne = Fichier.ReadLine
            If Ligne = "[MEMOIRE]" Then
                Ligne = Fichier.ReadLine
                Ligne = Fichier.ReadLine
                If InStr(Ligne, "Taille") > 0 Then Taille = Format(Val(Ligne.Substring(InStr(Ligne, "="))) / 1024, "# ### ###.00 Mo")
                Exit Do
            End If
        Loop Until Ligne Is Nothing
        Retour = Requete(ChaineDeConnexion, "UPDATE Serveur SET MEMOIRE_Taille='" & Taille & "' WHERE  POSTE_NomPoste)='" & CléServeur & "';")


        N = -1
        Do
            Ligne = Fichier.ReadLine
            If Ligne = "[HDD]" Then
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
                    Ret = Requete(ChaineDeConnexion, "INSERT INTO SERVEUR_HDD (HDD_Serveur, HDD_Lecteur, HDD_NS, HDD_Type, HDD_SysFic, HDD_EspLibre, HDD_EspTotal) values ('" &
                            CléServeur & "_" & Lecteur & "', '" & Lecteur & "', '" & NS & "', '" & Tipe & "', '" & SysFic & "', '" & EspLibre & "', '" & EspTotal & "')")
                    If Ret <> "" Then
                        Ret = Requete(ChaineDeConnexion, "UPDATE SERVEUR_HDD SET HDD_Serveur='" & CléServeur & "_" & Lecteur & "', HDD_Lecteur='" &
                            Lecteur & "', HDD_NS='" & NS & "', HDD_Type='" & Tipe & "', HDD_SysFic='" & SysFic & "', HDD_EspLibre='" & EspLibre & "', HDD_EspTotal='" & EspTotal & "' WHERE HDD_Serveur='" & CléServeur & "_" & Lecteur & "';")
                    End If
                Loop
                Exit Do
            End If
        Loop Until Ligne Is Nothing
        N = -1
        MAC = "temp"
        Do
            Ligne = Fichier.ReadLine
            If InStr(Ligne, "NomCarte") <> 0 Then MAC = "temp" : Ret = Requete(ChaineDeConnexion, "INSERT INTO SERVEUR_RESEAU (RESEAU_Serveur, RESEAU_NomCarte) VALUES ('" & CléServeur & "_" & MAC & "','" & Ligne.Substring(InStr(Ligne, "=")) & "')")
            If Ret <> "" Then Requete(ChaineDeConnexion, "UPDATE SERVEUR_Reseau SET RESEAU_NomCarte='" & Ligne.Substring(InStr(Ligne, "=")) & "' WHERE RESEAU_Serveur='" & CléServeur & "_" & MAC & "'") : Ret = ""
            If InStr(Ligne, "TypeCarte") <> 0 Then Requete(ChaineDeConnexion, "UPDATE SERVEUR_Reseau SET RESEAU_TypeCarte='" & Ligne.Substring(InStr(Ligne, "=")) & "' WHERE RESEAU_Serveur='" & CléServeur & "_" & MAC & "'")
            If InStr(Ligne, "Description") <> 0 Then Requete(ChaineDeConnexion, "UPDATE SERVEUR_Reseau SET RESEAU_Description='" & Ligne.Substring(InStr(Ligne, "=")) & "' WHERE RESEAU_Serveur='" & CléServeur & "_" & MAC & "'")
            If InStr(Ligne, "@MAC") <> 0 Then
                MAC = Ligne.Substring(InStr(Ligne, "="))
                Requete(ChaineDeConnexion, "UPDATE SERVEUR_Reseau SET RESEAU_MAC='" & MAC & "' WHERE RESEAU_Serveur='" & CléServeur & "_temp'")
                Requete(ChaineDeConnexion, "UPDATE SERVEUR_Reseau SET RESEAU_Serveur='" & CléServeur & "_" & MAC & "' WHERE RESEAU_Serveur='" & CléServeur & "_temp'")
            End If
            If InStr(Ligne, "VitesseMAX") <> 0 Then Requete(ChaineDeConnexion, "UPDATE SERVEUR_Reseau SET RESEAU_VitesseMAX='" & Ligne.Substring(InStr(Ligne, "=")) & "' WHERE RESEAU_Serveur='" & CléServeur & "_" & MAC & "'")
            If InStr(Ligne, "@IP") <> 0 Then
                If Ligne.Substring(InStr(Ligne, "=")) = "" Then
                    N = -1
                Else
                    N += 1
                    ADDIP = Ligne.Substring(InStr(Ligne, "=")) & "|" & ADDIP
                    Requete(ChaineDeConnexion, "UPDATE SERVEUR_Reseau SET RESEAU_IP='" & ADDIP.Substring(0, Len(ADDIP) - 1) & "' WHERE RESEAU_Serveur='" & CléServeur & "_" & MAC & "'")
                End If
            End If
            If InStr(Ligne, "MSR") <> 0 Then Requete(ChaineDeConnexion, "UPDATE SERVEUR_Reseau SET RESEAU_MSR='" & Ligne.Substring(InStr(Ligne, "=")) & "' WHERE RESEAU_Serveur='" & CléServeur & "_" & MAC & "'")
            If InStr(Ligne, "DHCP") <> 0 And InStr(Ligne, "@") = 0 Then Requete(ChaineDeConnexion, "UPDATE SERVEUR_Reseau SET RESEAU_DHCP='" & Ligne.Substring(InStr(Ligne, "=")) & "' WHERE RESEAU_Serveur='" & CléServeur & "_" & MAC & "'")
            If InStr(Ligne, "@DHCP") <> 0 Then Requete(ChaineDeConnexion, "UPDATE SERVEUR_Reseau SET RESEAU_ADDDHCP='" & Ligne.Substring(InStr(Ligne, "=")) & "' WHERE RESEAU_Serveur='" & CléServeur & "_" & MAC & "'")
            If InStr(Ligne, "@DNS") <> 0 Then Requete(ChaineDeConnexion, "UPDATE SERVEUR_Reseau SET RESEAU_DNS='" & Ligne.Substring(InStr(Ligne, "=")) & "' WHERE RESEAU_Serveur='" & CléServeur & "_" & MAC & "'")
            If Ligne = "[UTILISATEURS]" Then Exit Do
        Loop
        Requete(ChaineDeConnexion, "DELETE FROM SERVEUR_Reseau WHERE RESEAU_Serveur = '" & CléServeur & "_temp'")
        N = 0
        Do
            Ligne = Fichier.ReadLine
            If Ligne <> "" And Ligne <> "[GROUPES]" Then Utilisateurs = Utilisateurs & "|" & Ligne
        Loop While Not Ligne = "[GROUPES]"
        Requete(ChaineDeConnexion, "UPDATE Serveur SET UTILISATEURS='" & Utilisateurs & "' WHERE POSTE_NomPoste='" & CléServeur & "'")

        N = 0
        Do
            Ligne = Fichier.ReadLine
            If Ligne <> "" And Ligne <> "[STRATEGIE]" Then Groupes = Groupes & "|" & Ligne
            Groupes = Replace(Replace(Groupes, "�", ""), "'", " ")
        Loop While Not Ligne = "[STRATEGIE]"
        Requete(ChaineDeConnexion, "UPDATE Serveur SET GROUPES='" & Groupes & "' WHERE POSTE_NomPoste='" & CléServeur & "'")

        Do
            If Ligne = "[STRATEGIE]" Then
                Ligne = Fichier.ReadLine
                Ligne = Fichier.ReadLine
                If InStr(Ligne, "Expiration") > 0 Then Requete(ChaineDeConnexion, "UPDATE Serveur SET STRATEGIE_Expiration='" & Ligne.Substring(InStr(Ligne, "=")) & "' WHERE POSTE_NomPoste='" & CléServeur & "'")
                Ligne = Fichier.ReadLine
                If InStr(Ligne, "MDPVieMin") > 0 Then Requete(ChaineDeConnexion, "UPDATE Serveur SET STRATEGIE_MDPVieMin='" & Ligne.Substring(InStr(Ligne, "=")) & "' WHERE POSTE_NomPoste='" & CléServeur & "'")
                Ligne = Fichier.ReadLine
                If InStr(Ligne, "MDPVieMax") > 0 Then Requete(ChaineDeConnexion, "UPDATE Serveur SET STRATEGIE_MDPVieMax='" & Ligne.Substring(InStr(Ligne, "=")) & "' WHERE POSTE_NomPoste='" & CléServeur & "'")
                Ligne = Fichier.ReadLine
                If InStr(Ligne, "MDPLongueur") > 0 Then Requete(ChaineDeConnexion, "UPDATE Serveur SET STRATEGIE_MDPLongueur='" & Ligne.Substring(InStr(Ligne, "=")) & "' WHERE POSTE_NomPoste='" & CléServeur & "'")
                Ligne = Fichier.ReadLine
                If InStr(Ligne, "MDPAnterieur") > 0 Then Requete(ChaineDeConnexion, "UPDATE Serveur SET STRATEGIE_MDPAnterieur='" & Ligne.Substring(InStr(Ligne, "=")) & "' WHERE POSTE_NomPoste='" & CléServeur & "'")
                Ligne = Fichier.ReadLine
                If InStr(Ligne, "SeuilVerrou") > 0 Then Requete(ChaineDeConnexion, "UPDATE Serveur SET STRATEGIE_SeuilVerrou='" & Ligne.Substring(InStr(Ligne, "=")) & "' WHERE POSTE_NomPoste='" & CléServeur & "'")
                Ligne = Fichier.ReadLine
                If InStr(Ligne, "DureeVerrou") > 0 Then Requete(ChaineDeConnexion, "UPDATE Serveur SET STRATEGIE_DureeVerrou='" & Ligne.Substring(InStr(Ligne, "=")) & "' WHERE POSTE_NomPoste='" & CléServeur & "'")
                Ligne = Fichier.ReadLine
                If InStr(Ligne, "FenObsVerrou") > 0 Then Requete(ChaineDeConnexion, "UPDATE Serveur SET STRATEGIE_FenObsVerrou='" & Ligne.Substring(InStr(Ligne, "=")) & "' WHERE POSTE_NomPoste='" & CléServeur & "'")
                Ligne = Fichier.ReadLine
                If InStr(Ligne, "RolePoste") > 0 Then Requete(ChaineDeConnexion, "UPDATE Serveur SET STRATEGIE_RolePoste='" & Ligne.Substring(InStr(Ligne, "=")) & "' WHERE POSTE_NomPoste='" & CléServeur & "'")
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
        Requete(ChaineDeConnexion, "UPDATE Serveur SET Logiciels='" & Logiciels & "' WHERE POSTE_NomPoste='" & CléServeur & "'")

        N = 0
        Do
            Ligne = Fichier.ReadLine
            If Ligne <> "" And Ligne <> "[SERVICES]" Then N += 1 : Pilotes = Pilotes & "|" & Ligne
            Pilotes = Replace(Replace(Pilotes, "�", ""), "'", " ")
        Loop While Not Ligne = "[SERVICES]"
        Requete(ChaineDeConnexion, "UPDATE Serveur SET Pilotes='" & Pilotes & "' WHERE POSTE_NomPoste='" & CléServeur & "'")

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
                    service(1) = Nom.Substring(4)
                    service(2) = Replace(Description.Substring(12), "'", " ")
                    service(3) = Statut.Substring(7)
                    service(4) = Etat.Substring(8)
                    service(5) = CodeSortie.Substring(11)
                    Ligne = Fichier.ReadLine
                    Ret = Requete(ChaineDeConnexion, "INSERT INTO SERVEUR_SERVICES (SERVICES_Poste,SERVICES_Nom, SERVICES_Description, SERVICES_Statut, SERVICES_Etat, SERVICES_code) values ('" &
                            NomPoste & "_" & service(1) & "', '" & service(1) & "', '" & service(2) & "', '" & service(3) & "', '" & service(4) & "', '" & service(5) & "')")
                    If Ret <> "" Then
                        Ret = Requete(ChaineDeConnexion, "UPDATE SERVEUR_SERVICES SET SERVICES_Nom='" &
                            service(1) & "', SERVICES_Description='" & service(2) & "', SERVICES_Statut='" & service(3) & "', SERVICES_Etat='" & service(4) & "', SERVICES_code='" & service(5) & "' WHERE SERVICES_Poste='" & NomPoste & "_" & service(1) & "';")
                    End If
                Loop
                Exit Do
            End If
        Loop Until Ligne Is Nothing

        ' Création du fichier schéma pour l'importation du fichier texte dans la base
        Dim Schema As New StreamWriter("c:\temp\Schema.ini")
        Schema.WriteLine("[import.txt]")
        Schema.WriteLine("Format = Delimited(;)")
        Schema.WriteLine("ColNameHeader = True")
        Schema.Close()

        Ligne = Fichier.ReadLine
        Ligne = Fichier.ReadLine
        Ligne = Fichier.ReadLine
        Dim Import As New StreamWriter("c:\temp\import.txt")
        Import.WriteLine("MAJ_Nom;MAJ_Poste;MAJ_Descrption;MAJ_ID;MAJ_InstallPar;MAJ_Date")
        Do
            Ligne = Fichier.ReadLine
            If Ligne = "" Then Exit Do
            ' Recompilation au format CSV du fichier import.txt
            If Ligne <> "" Then
                MAJLigne(2) = ""
                Do
                    N += 1
                    MAJLigne(1) = Ligne.Substring(0, InStr(Ligne, "  ") - 1) ' Récupère le premier groupe de mot de ligne
                    MAJLigne(2) = MAJLigne(2) & ";" & MAJLigne(1) ' Ajoute le premier groupe à la chaine finale en y ajoutant un ;
                    Ligne = LTrim(Ligne.Substring(InStr(Ligne, "  ") - 1)) ' Supprime les espaces entre le début de la chaine et le groupe suivant
                Loop While InStr(Ligne, " ") > 0
                Import.WriteLine(MAJLigne(2).Substring(1)) ' Inscrit la chaine finale en supprimant le premier caratère puisque c'est un ;
            End If
        Loop
        Import.Close()

        Res = Requete(ChaineDeConnexion, "SELECT * INTO TEMP_MAJ FROM [Text;DATABASE=c:\temp;].[import.txt]") ' Importation du fichier import.txt dans la base de données en créant la table TEMP_MAJ
        If Res <> "" Then
            MsgBox("Une erreur c'est produite dans la base de données lors de l'ajout des MAJ : " & vbCrLf & Res)
            Res = Requete(ChaineDeConnexion, "DROP TABLE TEMP_MAJ") 'Suppression de la table TEMP_MAJ
        End If
        Res = Requete(ChaineDeConnexion, "DELETE FROM Serveur_MAJ WHERE MAJ_Poste='" & NomPoste & "';") 'Suppression de toutes les entrées dans la table Serveur_MAJ concernat ce serveur
        If Res <> "" Then MsgBox("Une erreur c'est produite dans la base de données lors de l'ajout des MAJ : " & vbCrLf & Res)
        Res = Requete(ChaineDeConnexion, "insert into Serveur_MAJ select * from TEMP_MAJ;") 'Copie des données entre table TEMP_MAJ et Serveur_MAJ
        If Res <> "" Then MsgBox("Une erreur c'est produite dans la base de données lors de l'ajout des MAJ : " & vbCrLf & Res)
        Res = Requete(ChaineDeConnexion, "DROP TABLE TEMP_MAJ") 'Suppression de la table TEMP_MAJ
        If Res <> "" Then MsgBox("Une erreur c'est produite dans la base de données lors de l'ajout des MAJ : " & vbCrLf & Res)
    End Sub
    Public Function CompactAndRepair(ByVal dbPath As String, ByVal tmpFolder As String) As Boolean
        'Declaration
        Dim oParams() As Object
        Dim oJRO As Object = Activator.CreateInstance(Type.GetTypeFromProgID("JRO.JetEngine"))
        Dim strLdbPath As String

        'Checked that database exist
        If Not File.Exists(dbPath) Then
            'File not found
            Return False
        End If

        'Checked that database is not open
        strLdbPath = Mid(dbPath, 1, Len(dbPath) - 3) & "ldb"
        If File.Exists(strLdbPath) Then
            'Database is open
            Return False
        End If

        'Compact and repair
        Try
            oParams = New Object() {"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbPath & ";Jet OLEDB:Engine Type=5", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & tmpFolder & "~tempdb.mdb;Jet OLEDB:Engine Type=5"}
            oJRO.GetType().InvokeMember("CompactDatabase", Reflection.BindingFlags.InvokeMethod, Nothing, oJRO, oParams)
            File.Delete(dbPath)
            File.Move(tmpFolder & "~tempdb.mdb", dbPath)

            Runtime.InteropServices.Marshal.ReleaseComObject(oJRO)
            oJRO = Nothing
            Return True

        Catch ex As Exception
            Return False
        End Try
    End Function
End Class
