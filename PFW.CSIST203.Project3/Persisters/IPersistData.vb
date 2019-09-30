Namespace PFW.CSIST203.Project3.Persisters

    ''' <summary>
    ''' Generic interface used for persisting data to and from a data source
    ''' </summary>
    Public Interface IPersistData
        Inherits IDisposable

        ''' <summary>
        ''' Retrieves a specific row number from the data source using a unique ID
        ''' </summary>
        ''' <param name="id">The unique identifier used by this persister to retrieve specific rows</param>
        ''' <returns>The data row representing the requested data</returns>
        Function GetRow(id As Integer) As System.Data.DataRow

        ''' <summary>
        ''' Retrieves a count of the number of elements present in the data source
        ''' </summary>
        ''' <returns>The number of items present in the underlying data source</returns>
        Function CountRows() As Integer

        ''' <summary>
        ''' Retrieves all data belonging to the data source as a .NET DataTable
        ''' </summary>
        ''' <returns>An in-memory representation of the data source's data as a data table</returns>
        Function GetData() As DataTable

        ''' <summary>
        ''' Updates the data row belonging to the underlying data source
        ''' </summary>
        ''' <param name="row">The data row that should be updated with all modifications made</param>
        Sub StoreRow(row As System.Data.DataRow)

        ''' <summary>
        ''' The file filter used by this persister when browsing for a file
        ''' </summary>
        ''' <returns>A valid string used by the OpenDialog browsing control</returns>
        ReadOnly Property FileFilter As String

        ''' <summary>
        ''' Create a DataRow compatible with the specified table
        ''' </summary>
        ''' <remarks>
        ''' NOTE: This method only creates a data row object, it does not change the database in any way
        ''' </remarks>
        ''' <param name="tableName">The table you wish to create a new row object for</param>
        ''' <returns>A DataRow initialized with the correct columns</returns>
        Function CreateRow(tableName As String) As DataRow


    End Interface

End Namespace


