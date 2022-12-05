Imports System.Data.SQLite
Imports System.Data

Public Class AccesDb
    Private _configDb As String = "samjin2DB.db"
    Private _connectionString As String = "Data Source={0};Version=3;"
    Public Sub New()
        _connectionString = String.Format(format:=_connectionString, arg0:=_configDb)
    End Sub
    Public Function QueryDL(Query As String) As ArrayList
        Dim SQL As String = String.Empty
        Dim dataList As New ArrayList
        Using conn As New SQLiteConnection(_connectionString)
            conn.Open()
            Using transaksi As SQLiteTransaction = conn.BeginTransaction()
                Using transaksi
                    Using cmd1 As New SQLiteCommand(conn)
                        Using cmd1
                            cmd1.Transaction = transaksi
                            SQL = Query
                            Debug.WriteLine("Query : " + SQL)
                            cmd1.CommandText = SQL
                            Using baca = cmd1.ExecuteReader()
                                While baca.Read()
                                    Try
                                        Dim unused = dataList.Add(baca.GetString(0))
                                        'Console.WriteLine(baca.GetString(0))
                                    Catch ex As Exception
                                        'Console.WriteLine("Get Code Serial")
                                        Dim unused1 = dataList.Add("0")
                                    End Try
                                End While
                            End Using
                        End Using
                    End Using
                    transaksi.Commit()
                End Using
            End Using
            conn.Close()
        End Using
        Return dataList
    End Function
    'var dataReader = cmd.ExecuteReader();
    'var dataTable = New DataTable();
    'dataTable.Load(dataReader);
    Public Function QueryDLtable(Query As String) As DataTable
        Dim SQL As String = String.Empty
        'DataTable.EnforceConstraints.Set(False)
        Dim dataSet As New DataSet("dataSet")
        Dim dt As New DataTable("table")
        dataSet.Tables.Add(dt)
        dataSet.EnforceConstraints = False
        Using conn As New SQLiteConnection(_connectionString)
            conn.Open()
            Using transaksi As SQLiteTransaction = conn.BeginTransaction()
                Using transaksi
                    Using cmd1 As New SQLiteCommand(conn)
                        Using cmd1
                            cmd1.Transaction = transaksi
                            SQL = Query
                            Debug.WriteLine("Query : " + SQL)
                            cmd1.CommandText = SQL
                            Using baca = cmd1.ExecuteReader()
                                'dt.e
                                dt.Load(baca)
                            End Using
                        End Using
                    End Using
                    transaksi.Commit()
                End Using
            End Using
            conn.Close()
        End Using
        Return dt
    End Function
    Public Sub QueryUL(Query As String)
        Dim dataList As ArrayList = New ArrayList
        Using conn As New SQLiteConnection(connectionString:=_connectionString)
            conn.Open()
            Dim transaksi As SQLiteTransaction = conn.BeginTransaction()
            Using transaksi
                Using cmd1 As New SQLiteCommand(conn)
                    Using cmd1
                        cmd1.Transaction = transaksi
                        Dim SQL As String = Query
                        cmd1.CommandText = SQL
                        Dim unused = cmd1.ExecuteNonQuery()
                        Console.WriteLine("Uploaded")
                    End Using
                End Using
                transaksi.Commit()
            End Using
            conn.Close()
        End Using
    End Sub
End Class