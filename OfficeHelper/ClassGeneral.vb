Imports System.Data.SqlClient
Public Class ClassGeneral


    Dim cn As New SqlConnection

    Public Function AbrirConexionSql() As Boolean
        cn = New SqlConnection("Data Source=SERVIDORPG\SQL;Initial Catalog=VENTAS;User ID=iEx;Password=iEx;MultipleActiveResultSets=True")

        Try
            If cn.State = ConnectionState.Closed Then
                cn.Open()
                Return True
            End If
        Catch ex As Exception
            Return False
        End Try

    End Function

    Public Function tblClientes() As DataTable
        If AbrirConexionSql() = True Then

            Dim tbl As New DSGeneral.tblClienteDataTable

            Dim cmd As New SqlCommand("SELECT Clientes.NITCLIE, Clientes.NOMCLIE, Clientes.DIRCLIE, Municipios.DESMUNI, Departamentos.DESDEPTO
                                       FROM Clientes LEFT OUTER JOIN Departamentos ON Clientes.CODDEPTO = Departamentos.CODDEPTO AND Clientes.EMP_NIT = Departamentos.EMP_NIT LEFT OUTER JOIN
                                       Municipios ON Clientes.CODMUNI = Municipios.CODMUNI AND Clientes.EMP_NIT = Municipios.EMP_NIT", cn)
            Dim dr As SqlDataReader = cmd.ExecuteReader
            dr.Read()
            Do While dr.Read
                tbl.Rows.Add(New Object() {
                        dr(0),
                        dr(1),
                        dr(2),
                        dr(3),
                        dr(4)
                })
            Loop
            dr.Close()
            cmd.Dispose()

            Return tbl
        Else
            MsgBox("No se pudo abrir la conexión")
        End If

    End Function



End Class
