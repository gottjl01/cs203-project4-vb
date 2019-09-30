Namespace PFW.CSIST203.Project3.Persisters.Access

    ''' <summary>
    ''' Persister that interacts with an Access database
    ''' </summary>
    Public Class AccessPersister
        Implements IPersistData

        Friend ReadOnly accessFile As String = Nothing
        Friend noDatabase As Boolean = False
        Friend IsDisposed As Boolean = False

        ''' <summary>
        ''' Creates a persister that acts like no data exists
        ''' </summary>
        Public Sub New()
            ' create an access persister that does nothing
            noDatabase = True
        End Sub

        ''' <summary>
        ''' Creates a persister that use the supplied access database file as its source
        ''' </summary>
        ''' <param name="accessFile">The access database file to read</param>
        Public Sub New(accessFile As String)
            If Not System.IO.File.Exists(accessFile) Then
                Throw New System.IO.FileNotFoundException("Access Database not found", accessFile)
            End If
            Me.accessFile = accessFile
        End Sub

        ''' <summary>
        ''' The filter used by the open dialog to find files that this persister will handle
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property FileFilter As String Implements IPersistData.FileFilter
            Get
                Return "Access Database Files|*.mdb;*.accdb"
            End Get
        End Property

        ''' <summary>
        ''' Retrieves a row from the access database using the specific ID
        ''' </summary>
        ''' <param name="id">The ID of the row to retrieve</param>
        ''' <returns>A DataRow representing the retrieved data, or null (Nothing) it is not found</returns>
        Public Function GetRow(id As Integer) As DataRow Implements IPersistData.GetRow

            If IsDisposed Then
                Throw New ObjectDisposedException("Persister")
            End If

            If noDatabase Then
                Return Nothing
            End If
            Dim dt As New DataTable("tblEmployees")
            Using connection As New System.Data.OleDb.OleDbConnection(Util.GetAccessConnectionString(accessFile))
                connection.Open()
                Using cmd = connection.CreateCommand()
                    cmd.CommandText = "SELECT * FROM [tblEmployees] where [ID] = @ID"

                    ' create the single parameter object
                    Dim par = cmd.CreateParameter()
                    par.ParameterName = "@ID"
                    par.DbType = DbType.Int32
                    par.Value = id
                    cmd.Parameters.Add(par)

                    Using dr = cmd.ExecuteReader()
                        dt.Load(dr)
                    End Using

                End Using
            End Using

            Dim row As System.Data.DataRow = Nothing

            ' single the single row that was returned if exactly one match was found
            If dt.Rows.Count = 1 Then
                row = dt.Rows(0)
            End If

            Return row

        End Function

        ''' <summary>
        ''' Counts the number of rows present in the access database
        ''' </summary>
        ''' <returns>The number of rows present in the database</returns>
        Public Function CountRows() As Integer Implements IPersistData.CountRows

            If IsDisposed Then
                Throw New ObjectDisposedException("Persister")
            End If

            If noDatabase Then
                Return 0
            End If

            Using connection As New System.Data.OleDb.OleDbConnection(Util.GetAccessConnectionString(accessFile))
                connection.Open()
                Using cmd = connection.CreateCommand()
                    cmd.CommandText = "SELECT count(*) FROM [tblEmployees]"
                    Return Integer.Parse(cmd.ExecuteScalar().ToString())
                End Using
            End Using

        End Function

        ''' <summary>
        ''' Disposes of the object and any managed resources
        ''' </summary>
        Public Sub Dispose() Implements IDisposable.Dispose
            IsDisposed = True
        End Sub

        ''' <summary>
        ''' Retrieves all data present in the access database as a single operation
        ''' </summary>
        ''' <returns>A datatable representing all data present in the access database</returns>
        Public Function GetData() As DataTable Implements IPersistData.GetData

            If IsDisposed Then
                Throw New ObjectDisposedException("Persister")
            End If

            Dim dt As New DataTable("tblEmployees")
            If noDatabase Then
                Return dt
            End If

            Using connection As New System.Data.OleDb.OleDbConnection(Util.GetAccessConnectionString(accessFile))
                connection.Open()
                Using cmd = connection.CreateCommand()
                    cmd.CommandText = "SELECT * FROM [tblEmployees]"

                    Using dr = cmd.ExecuteReader()
                        dt.Load(dr)
                    End Using

                End Using
            End Using
            Return dt
        End Function

        Public Sub StoreRow(row As DataRow) Implements IPersistData.StoreRow
            Throw New NotImplementedException()
        End Sub

        Public Function CreateRow(tableName As String) As DataRow Implements IPersistData.CreateRow
            Throw New NotImplementedException()
        End Function

    End Class

End Namespace

