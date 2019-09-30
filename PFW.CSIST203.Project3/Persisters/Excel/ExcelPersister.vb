Namespace PFW.CSIST203.Project3.Persisters.Excel

    ''' <summary>
    ''' Excel Persister that interacts with data in an xls or xlsx file
    ''' </summary>
    Public Class ExcelPersister
        Implements IPersistData

        Private logger As log4net.ILog = log4net.LogManager.GetLogger(GetType(ExcelPersister))

        Private _Data As System.Data.DataTable = Nothing
        Private _ExcelFile As String = Nothing

        ''' <summary>
        ''' This data table must be populated with all data contained in the specified excel file
        ''' </summary>
        Friend Property Data As System.Data.DataTable
            Get
                Return _Data
            End Get
            Private Set(value As System.Data.DataTable)
                _Data = value
            End Set
        End Property

        Private _isDisposed As Boolean = False

        ''' <summary>
        ''' Get a value indicating whether or not the object has been disposed
        ''' </summary>
        Friend Property isDisposed As Boolean
            Get
                Return _isDisposed
            End Get
            Private Set(value As Boolean)
                _isDisposed = value
            End Set
        End Property

        Public ReadOnly Property FileFilter As String Implements IPersistData.FileFilter
            Get
                Return "Excel Files|*.xls;*.xlsx"
            End Get
        End Property

        ''' <summary>
        ''' This contructor creates a persister that contains no data
        ''' </summary>
        Public Sub New()
            Data = New DataTable("Sheet1")
            Data.Columns.AddRange(
                {
                    New DataColumn("First Name", GetType(String)),
                    New DataColumn("Last Name", GetType(String)),
                    New DataColumn("E-mail Address", GetType(String)),
                    New DataColumn("Business Phone", GetType(String)),
                    New DataColumn("Company", GetType(String)),
                    New DataColumn("Job Title", GetType(String))
                })
        End Sub

        ''' <summary>
        ''' Persists data to and from the specified excel file
        ''' </summary>
        ''' <param name="excelFilepath">The excel file that should be read into memory</param>
        Public Sub New(excelFilepath As String)
            If Not System.IO.File.Exists(excelFilepath) Then
                Throw New System.IO.FileNotFoundException(excelFilepath)
            End If

            ' keep a pointer to the excel file
            _ExcelFile = excelFilepath

            Dim table = New System.Data.DataTable("Sheet1")
            Try
                Using connection As New System.Data.OleDb.OleDbConnection(Util.GetExcelConnectionString(excelFilepath, True))
                    Using cmd As System.Data.IDbCommand = connection.CreateCommand()
                        cmd.CommandText = "SELECT * FROM [Sheet1$]"
                        connection.Open()
                        Using dr As System.Data.IDataReader = cmd.ExecuteReader()
                            table.Load(dr)
                        End Using
                    End Using
                End Using

            Catch ex As Exception

                If (ex.Message.IndexOf("Microsoft.ACE.OLEDB.12.0' provider is not registered on the local machine", StringComparison.OrdinalIgnoreCase) >= 0) Then
                    logger.Error("Please install the Microsoft Access Database Engine 2010 Redistributable: https://www.microsoft.com/en-US/download/details.aspx?id=13255", ex)
                    Throw New System.Exception("Please install the Microsoft Access Database Engine 2010 Redistributable: https://www.microsoft.com/en-US/download/details.aspx?id=13255")
                End If

                logger.Error("Problem reading excel file: " + excelFilepath, ex)
                Throw
            End Try

            ' assign the data table before exiting the constructor
            Me.Data = table

        End Sub

        ''' <summary>
        ''' This method must retrieve data from the data table at the specified row
        ''' </summary>
        ''' <param name="rowNumber">The row you would like to request from the excel data (zero based)</param>
        ''' <returns>The row requested or nothing if the row number specified is non-positive or beyond the end of the available data</returns>
        Public Function GetRow(rowNumber As Integer) As System.Data.DataRow Implements IPersistData.GetRow
            If Me.isDisposed Then
                Throw New ObjectDisposedException("Data (DataTable)")
            End If
            If rowNumber < 0 OrElse rowNumber >= Me.Data.Rows.Count Then
                Return Nothing
            End If
            Return Me.Data.Rows(rowNumber)
        End Function

        ''' <summary>
        ''' Retrieves the number of rows contained in the excel data
        ''' </summary>
        Public Function CountRows() As Integer Implements IPersistData.CountRows
            If Me.isDisposed Then
                Throw New ObjectDisposedException("Data (DataTable)")
            End If
            Return Me.Data.Rows.Count
        End Function

        ''' <summary>
        ''' Cleans up any managed resources used by the DataTable
        ''' </summary>
        Public Sub Dispose() Implements IPersistData.Dispose
            isDisposed = True
            If Not Data Is Nothing Then
                Data.Dispose()
                Data = Nothing
            End If
        End Sub

        Public Function GetData() As DataTable Implements IPersistData.GetData
            Return Data
        End Function

        Public Sub StoreRow(row As DataRow) Implements IPersistData.StoreRow
            Throw New NotImplementedException()
        End Sub

        Public Function CreateRow(tableName As String) As DataRow Implements IPersistData.CreateRow
            Throw New NotImplementedException()
        End Function

    End Class

End Namespace


