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
'    Foundation, Inc., 675 Mass Ave, Cambridge, MA 02139, USA".Public Class Form2

Imports System.Data.SqlClient
Imports System.Data.SqlServerCe

Public Class Form4
    Private Sub Form4_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            Dim SQLConn = New SqlConnection(), cmd As String
            SQLConn.ConnectionString = "Data Source=(LocalDb)\v11.0;AttachDbFilename=|DataDirectory|\BDD.sdf;Initial Catalog=Ping"

            SQLConn.Open()

            cmd = "INSERT INTO MaTable (ID, Action) VALUES ('TEST', 'TEST')"

            Dim Command = New SqlCommand(cmd, SQLConn)
            Command.ExecuteNonQuery()
            SQLConn.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
            End

        End Try
    End Sub
End Class