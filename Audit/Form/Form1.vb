

Imports System.Data.OleDb

Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim cmd As New OleDbCommand, id
        Try
            id = Open()
            Label1.Text = "Connection OK"
            cmd.CommandText = "INSERT INTO Serveur
                                         (ID, [Action])
                                         VALUES        ('test456', 'test')"
            cmd.Connection = id
            cmd.ExecuteNonQuery()
        Catch ex As OleDbException
            MsgBox(ex.Message)
            End
            Close()
        Catch ex As InvalidOperationException
            MsgBox(ex.Message)
            End
        End Try
        Close()
    End Sub

    Public Const ChaineDeConnexion As String = "Provider=microsoft.jet.oledb.4.0;Data Source=D:\Users\u165147\source\repos\Audit\Audit\BDD\BDD.mdb"
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
End Class
