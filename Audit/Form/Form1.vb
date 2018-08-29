

Imports System.Data.OleDb

Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim NetSql As New SQL
        NetSql.ChargementFichier("c:\temp\Audit_NA1VM28_13082018.txt")
    End Sub

End Class
