' déclaration des variables
Option Explicit

'on error resume next 

Const strTypeComputer = "workstation"
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const HKEY_LOCAL_MACHINE = &H80000002
Const UnInstPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"

Dim objWMIService, objItem, colOperatingSystem, colBaseBoards, objProcessor, writetextfile, fso1, colSettings, Architecture
Dim oUnLecteur, strLectType, oDesLecteurs, oFSO, objLogicalDisk, colAdapters, colAdaptersConf, objAdapter, objAdapterConf
Dim pass, colUsagers, objUsager, Poste, strComputer, oShell, env, FichierTxt, shl, Fs, Depart, Memoire, ScreenHeight,ScreenWeight 
Dim Fin, User, S, strLine, ReadTextFile, N, X, Group, Domaine, colMemoryCapacity, colComputerSystem, Fabricant, Modele
Dim colEcrans, objEcran, colInstalledPrinters, objPrinter, colSoundDevices, objSoundDevice, colCartesVideo, objCartesVideo, A
Dim subkey, arrSubKeys, software, oReg, Thermal, objService, colService, MAJ, FSO2, strReportProc
Dim Valeur (10, 10)

Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
Set oShell = CreateObject("wscript.Shell")
Set env = oShell.environment("Process")
strComputer = env.Item("Computername")
strReportProc = "winaudit.txt"
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set oFSO = WScript.CreateObject("Scripting.FileSystemObject")
Set oDesLecteurs = oFSO.Drives
Set FSO1 = CreateObject("Scripting.FileSystemObject")
Set FSO2 = CreateObject("Scripting.FileSystemObject")
Set WriteTextFile = FSO1.OpenTextFile(strReportProc, ForWriting, True)


' Récupération des informations

 '=======================================POSTE=======================================

WriteTextFile.WriteLine "[POSTE]" & vbcrlf

Set colComputerSystem =  objWMIService.ExecQuery ("Select * from Win32_ComputerSystem")
For Each objItem in colComputerSystem
	Domaine = objItem.domain
	Memoire = objItem.TotalPhysicalMemory
	Fabricant = objItem.Manufacturer
	Modele = objItem.Model
	Thermal = objItem.ThermalState
	Select case Thermal
		case 1 = Thermal = "Autre"
		case 2 = Thermal = "Inconnu"
		case 3 = Thermal = "OK"
		case 4 = Thermal = "Alerte"
		case 5 = Thermal = "Critique"
		case 6 = Thermal = "Non réquapérable"
	End Select
Next

Set colOperatingSystem = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
For Each objItem in colOperatingSystem
	WriteTextFile.WriteLine "NomPoste=" & strComputer
	WriteTextFile.WriteLine "DescPoste=" & objItem.Description
	WriteTextFile.WriteLine "OS=" & objItem.Caption
	WriteTextFile.WriteLine "Version=" & objItem.Version
	WriteTextFile.WriteLine "DateInstall=" & objItem.InstallDate
	WriteTextFile.WriteLine "NumDernierSPMa=" & objItem.ServicePackMajorVersion
	WriteTextFile.WriteLine "NumDernierSPMi=" & objItem.ServicePackMinorVersion
	WriteTextFile.WriteLine "Fabricant=" & Fabricant
	WriteTextFile.WriteLine "Modele=" & Modele
Next

 '=======================================CARTE-MERE=======================================

WriteTextFile.WriteLine vbcrlf & "[CARTE-MERE]" & vbcrlf
Set colBaseBoards =  objWMIService.ExecQuery ("Select * from Win32_BaseBoard")
For Each objItem in colBaseBoards
	WriteTextFile.WriteLine "Nom=" & objItem.Name
	WriteTextFile.WriteLine "Modèle=" & objItem.Model
	WriteTextFile.WriteLine "Manufacturier=" & objItem.Manufacturer
Next

 '=======================================PROCESSEUR=======================================

WriteTextFile.WriteLine vbcrlf & "[PROCESSEUR]" & vbcrlf
Set colSettings = objWMIService.ExecQuery ("Select * from Win32_Processor")
For Each objProcessor in colSettings
        If objProcessor.Architecture = 0 Then
                Architecture = "x86"
        ElseIf objProcessor.Architecture = 1 Then
                Architecture = "MIPS"
        ElseIf objProcessor.Architecture = 2 Then
                Architecture = "Alpha"
        ElseIf objProcessor.Architecture = 3 Then
                Architecture = "PowerPC"
        ElseIf objProcessor.Architecture = 6 Then
                Architecture = "ia64"
        Else
                Architecture = "inconnu"
        End If
		WriteTextFile.WriteLine "TypeProc=" & Architecture
        WriteTextFile.WriteLine "NomProc=" & objProcessor.Name
        WriteTextFile.WriteLine "DescProc=" & objProcessor.Description
        WriteTextFile.WriteLine "VitesseACT=" & objProcessor.CurrentClockSpeed & " Mhz"
        WriteTextFile.WriteLine "VitesseMAX=" & objProcessor.MaxClockSpeed & " Mhz"
Next

 '=======================================MEMOIRE=======================================

WriteTextFile.WriteLine vbcrlf & "[MEMOIRE]" & vbcrlf
WriteTextFile.WriteLine "Taille=" & Memoire

 '=======================================HDD=======================================

WriteTextFile.WriteLine vbcrlf & "[HDD]" & vbcrlf
For Each oUnLecteur in oDesLecteurs
        If oUnLecteur.IsReady Then
				If oUnLecteur.DriveType = 0 Then strLectType = "Inconnu"
				If oUnLecteur.DriveType = 1 Then strLectType = "Amovible (Disquette, clé USB, etc.)"
				If oUnLecteur.DriveType = 2 Then strLectType = "Fixe (Disque dur, etc.)"
				If oUnLecteur.DriveType = 3 Then strLectType = "Réseau"
				If oUnLecteur.DriveType = 4 Then strLectType = "CD-Rom"
				If oUnLecteur.DriveType = 5 Then strLectType = "Virtuel"
 
                WriteTextFile.WriteLine "Lecteur=" & oUnLecteur.DriveLetter
                WriteTextFile.WriteLine "NS=" & oUnLecteur.SerialNumber
                WriteTextFile.WriteLine "Type=" & strLectType 
                WriteTextFile.WriteLine "SysFic=" & oUnLecteur.FileSystem
 
                Set objLogicalDisk = GetObject("winmgmts:Win32_LogicalDisk.DeviceID='" & oUnLecteur.DriveLetter & ":'")
                WriteTextFile.WriteLine "EspLibre=" & objLogicalDisk.FreeSpace /1024\1024+1
                WriteTextFile.WriteLine "EspTotal=" & objLogicalDisk.Size /1024\1024+1 & vbcrlf
 
        End If
Next

 '=======================================RESEAU=======================================

WriteTextFile.WriteLine vbcrlf & "[RESEAU]" & vbcrlf
Set colAdapters = objWMIService.ExecQuery ("Select * from Win32_NetworkAdapter")
	WriteTextFile.WriteLine "Domaine=" & Domaine
	For Each objAdapter in colAdapters
        If (objAdapter.Manufacturer<>"Microsoft") Then
                WriteTextFile.WriteLine "NomCarte=" & objAdapter.Name
                WriteTextFile.WriteLine "TypeCarte=" & objAdapter.AdapterType
                WriteTextFile.WriteLine "Description=" & objAdapter.Description
                WriteTextFile.WriteLine "@MAC=" & objAdapter.MACAddress
                WriteTextFile.WriteLine "VitesseMAX=" & objAdapter.MaxSpeed  & vbcrlf
 
                Set colAdaptersConf = objWMIService.ExecQuery ("Select * from Win32_NetworkAdapterConfiguration")
                For Each objAdapterConf in colAdaptersConf
                        If (objAdapter.Manufacturer<>"Microsoft") AND (objAdapter.Caption = objAdapterConf.Caption) Then
                                WriteTextFile.WriteLine "	@IP=" & objAdapterConf.IPAddress(0)
                                WriteTextFile.WriteLine "	MSR=" & objAdapterConf.IPSubnet(0)
                                WriteTextFile.WriteLine "	DHCP=" & objAdapterConf.DHCPEnabled
                                WriteTextFile.WriteLine "	@DHCP=" & objAdapterConf.DHCPServer
                                WriteTextFile.WriteLine "	@DNS=" & objAdapterConf.DNSServerSearchOrder(0) & vbcrlf
								pass = true
                        End If
						if (pass=true) then exit for 
                Next
        End If
Next

 '=======================================UTILISATEURS=======================================

WriteTextFile.WriteLine vbcrlf & "[UTILISATEURS]" 
'Fichier où seront copiées les données de la commande Dos
FichierTxt = "temp.txt"

set shl = createobject("wscript.shell")
shl.run "cmd /c Net User > " & FichierTxt, 0, true
Set Fs = CreateObject("Scripting.FileSystemObject")

'Délai pour laisser le temps de créer le fichier texte
WScript.Sleep (1000)

S = Fs.OpenTextFile(FichierTxt, 1).ReadAll
Depart = instr(1, S, "-------------------------------------------------------------------------------")+79
Fin = instr(1, S, "command")
User = mid(S, Depart, Fin-4-Depart)
User = replace(user, "‚", "é")
do while instr(1, User, " ")<> 0
	A = A & left(User, instr(1, User, " ")) & vbcrlf
	User = ltrim(right(User, len(User) - instr(1, User, " ")))
loop
A = A & user
WriteTextFile.WriteLine A

 '=======================================GROUPES=======================================

WriteTextFile.WriteLine vbcrlf & "[GROUPES]" 
'Fichier où seront copiées les données de la commande Dos
FichierTxt = "temp.txt"

set shl = createobject("wscript.shell")
shl.run "cmd /c Net localgroup > " & FichierTxt, 0, true

'Délai pour laisser le temps de créer le fichier texte
WScript.Sleep (1000)

S = Fs.OpenTextFile(FichierTxt, 1).ReadAll
Depart = instr(1, S, "-------------------------------------------------------------------------------")+79
Fin = instr(1, S, "command")
Group = mid(S, Depart, Fin-4-Depart)
Group = replace(Group, "‚", "é")
Group = replace(Group, "*", "")
WriteTextFile.WriteLine Group

 '=======================================STRATEGIE=======================================

WriteTextFile.WriteLine vbcrlf & "[STRATEGIE]" & vbcrlf
'Fichier où seront copiées les données de la commande Dos

shl.run "cmd /c Net accounts > " & FichierTxt, 0, true
Set ReadTextFile = FSO2.OpenTextFile(FichierTxt, ForReading,False)

'Délai pour laisser le temps de créer le fichier texte
WScript.Sleep (1000)

n = 0

' Définition des clés
Valeur(1, 0) = "Expiration"
Valeur(2, 0) = "MDPVieMin"
Valeur(3, 0) = "MDPVieMax"
Valeur(4, 0) = "MDPLongueur"
Valeur(5, 0) = "MDPAnterieur"
Valeur(6, 0) = "SeuilVerrou"
Valeur(7, 0) = "DureeVerrou"
Valeur(8, 0) = "FenObsVerrou"
Valeur(9, 0) = "RolePoste"

Do Until ReadTextFile.AtEndOfStream
	n = n + 1
    Valeur(n, 1) = ReadTextFile.ReadLine
	Valeur(n, 1) = lTrim(Rtrim(right(Valeur(n, 1), len(Valeur(n, 1)) - instr(1, Valeur(n, 1), ":")-1)))
	'wscript.echo Valeur(n, 1)
	if n=10 then exit do
Loop

for X=1 to 9
	WriteTextFile.WriteLine Valeur(X, 0) & "=" & Valeur(X, 1)
next	

'=======================================LOGICIELS=======================================
WriteTextFile.WriteLine vbcrlf & "[LOGICIELS]" & vbcrlf
oReg.EnumKey HKEY_LOCAL_MACHINE, UnInstPath, arrSubKeys
For Each subkey In arrSubKeys
        'MsgBox subkey
        If Left (subkey, 1) <> "{" Then
                software = software & subkey & vbCrLf
        End If
Next
WriteTextFile.WriteLine software

'=======================================PILOTES=======================================

WriteTextFile.WriteLine vbcrlf & "[PILOTES]" & vbcrlf
Set colCartesVideo = objWMIService.ExecQuery ("Select Description From Win32_VideoController") 
For Each objCartesVideo in colCartesVideo
        WriteTextFile.WriteLine "Video=" & objCartesVideo.Description 
Next
WriteTextFile.WriteLine vbcrlf
Set colSoundDevices = objWMIService.ExecQuery ("Select Description From Win32_SoundDevice") 
For Each objSoundDevice in colSoundDevices
        WriteTextFile.WriteLine "Son=" & objSoundDevice.Description 
Next
WriteTextFile.WriteLine vbcrlf
Set colInstalledPrinters =  objWMIService.ExecQuery ("Select * from Win32_Printer")
	For Each objPrinter  in colInstalledPrinters 
        WriteTextFile.WriteLine "Imprimante=" & objPrinter.Name  
Next
WriteTextFile.WriteLine vbcrlf
Set colEcrans =  objWMIService.ExecQuery ("Select * from Win32_DesktopMonitor")
	For Each objEcran in colEcrans
        WriteTextFile.WriteLine "Moniteur=" & objEcran.Name  
        WriteTextFile.WriteLine "	Type=" & objEcran.MonitorType  
        WriteTextFile.WriteLine "	Fabricant=" & objEcran.MonitorManufacturer 
        WriteTextFile.WriteLine "	Hauteur=" & objEcran.ScreenHeight  
        WriteTextFile.WriteLine "	Largeur=" & objEcran.ScreenWidth  & vbcrlf 
Next

 '=======================================SERVICES=======================================

WriteTextFile.WriteLine vbcrlf & "[SERVICES]" & vbcrlf
Set colService =  objWMIService.ExecQuery ("Select * from Win32_Service")
	For Each objService in colService
        WriteTextFile.WriteLine "Nom=" & objService.Name  
        WriteTextFile.WriteLine "Description=" & objService.Description  
        WriteTextFile.WriteLine "Statut=" & objService.Status 
        WriteTextFile.WriteLine "Demarre=" & objService.Started  
        WriteTextFile.WriteLine "CodeSortie=" & objService.ExitCode   & vbcrlf
Next

 '=======================================MAJ=======================================

WriteTextFile.WriteLine vbcrlf & "[MAJ]" & vbcrlf
shl.run "cmd /c wmic qfe > " & FichierTxt, 0, true
Set ReadTextFile = FSO2.OpenTextFile(FichierTxt, ForReading,False)
WScript.Sleep (20000)

WriteTextFile.close

shl.run "cmd /c type temp.txt >> " & strReportProc
WScript.Sleep (500)
shl.run "cmd /c del temp.txt"

