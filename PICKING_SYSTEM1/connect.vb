Imports System.Runtime.InteropServices
Imports System.Data
'Imports System.Data.SqlServerCe
Imports System.Data.SqlClient
Imports System.Configuration
Public Class connect
    Public Sub conn()
        Dim myConn As SqlConnection
        ' myConn = New SqlConnection("Data Source=192.168.10.16\SQLEXPRESS2017,1433;Initial Catalog=tbkkfa01_dev;Integrated Security=False;User Id=sa;Password=p@sswd;")
        myConn = New SqlConnection("Data Source=192.168.161.101;Initial Catalog=tbkkfa01_dev;Integrated Security=False;User Id=pcs_admin;Password=P@ss!fa")
        'Server= localhost; Database= employeedetails; Integrated Security=True;
        'Data Source=192.168.10.16\SQLEXPRESS,1433;Initial Catalog=tbkkfa01_dev;Integrated Security=SSPI;
        myConn.Open()
    End Sub
End Class
