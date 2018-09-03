

Imports System.IO
Imports System.Data.OleDb
Imports Microsoft.Office.Interop.Access

Public Class Form1
    Public Debut, Temps
    Public Const ChaineDeConnexion As String = "Provider=microsoft.jet.oledb.4.0;Data Source=D:\Users\u165147\source\repos\Audit\Audit\BDD\BDD.mdb"
    Declare Function GetTickCount Lib "kernel32" Alias "GetTickCount" () As Long
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim ThreadCollect As New System.Threading.Thread(AddressOf Collect)
        ThreadCollect.Start() ' Démarrer le nouveau thread.
    End Sub

    Private Sub NsTheme1_Click(sender As Object, e As EventArgs) Handles NsTheme1.Click

    End Sub

    Public Sub Collect()
        Debut = GetTickCount
        Dim NetSql As New SQL, Ligne As String, Ena(10) As String, Ret As String, RequeteQ As OleDb.OleDbDataReader
        Dim ServeurLst As New StreamReader("C:\varsoft\chksys\.enable_win.lst")
        Dim connect As New OleDbConnection(ChaineDeConnexion)
        Dim cmd As New OleDbCommand
        NsProgressBar1.Maximum = 1
        Do
            Ligne = ServeurLst.ReadLine
            If Ligne IsNot Nothing Then
                Ena = Ligne.Split(":")
                Try
                    Ret = NetSql.Requete(ChaineDeConnexion, "INSERT INTO ServeurLst (Etat, Nom, IP, DerniereFoisVu) VALUES ('" & Ena(0) & "', '" & Ena(1) & "', '" & Ena(2) & "', '" & Ena(5) & "')")
                    If Ret <> "" Then Ret = NetSql.Requete(ChaineDeConnexion, "UPDATE ServeurLst SET Etat='" & Ena(0) & "', Nom='" & Ena(1) & ", IP='" & Ena(2) & "', DerniereFoisVu='" & Ena(5) & "') WHERE Nom='" & Ena(1))
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
                NsProgressBar1.Maximum += 1
            End If
        Loop Until Ligne Is Nothing
        NsProgressBar1.Maximum -= 1
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

        'NetSql.ChargementFichier("c:\temp\Audit_NA1VM28_13082018.txt")
        NsLabel2.Value1 = (GetTickCount - Debut) / 60000
    End Sub

End Class
