Imports Microsoft.VisualStudio.TestTools.UnitTesting

Namespace PFW.CSIST203.Project3.Tests.Persisters.Access

    ''' <summary>
    ''' Root testing harness for the AccessPersister
    ''' </summary>
    Public MustInherit Class AccessPersisterTests
        Inherits TestBase

        ''' <summary>
        ''' Refactored method for creating a persister that uses the parameterless constructor
        ''' </summary>
        ''' <returns>A persister that should contain no data</returns>
        Protected Function CreatePersister() As PFW.CSIST203.Project3.Persisters.Access.AccessPersister

            Dim obj As PFW.CSIST203.Project3.Persisters.Access.AccessPersister

            AssertDelegateSuccess(
                Sub()
                    obj = New Project3.Persisters.Access.AccessPersister()
                End Sub,
                "Instantiation of the default constructor should not throw an exception")

            Return obj

        End Function

        ''' <summary>
        ''' Refactored method for creating a persister that uses the sample access database embedded into the project
        ''' </summary>
        ''' <param name="workingDirectory">The method specific working directory</param>
        ''' <returns>A persister instantiated using the sample access database</returns>
        Protected Function CreatePersister(workingDirectory As String) As PFW.CSIST203.Project3.Persisters.Access.AccessPersister

            Dim obj As PFW.CSIST203.Project3.Persisters.Access.AccessPersister
            Dim tmpAccessDatabaseFile = System.IO.Path.Combine(workingDirectory, "sample-data.accdb")
            CopyEmbeddedResourceBaseToDirectory("PFW.CSIST203.Project3.Tests.Resources.Data", workingDirectory)
            Assert.IsTrue(System.IO.File.Exists(tmpAccessDatabaseFile), "The sample data access database file was not found")

            AssertDelegateSuccess(
                Sub()
                    obj = New Project3.Persisters.Access.AccessPersister(tmpAccessDatabaseFile)
                End Sub,
                "Instantiation pointing to a file should not immediately throw an exception")

            Return obj

        End Function

        ''' <summary>
        ''' Constructor testing harness
        ''' </summary>
        <TestClass>
        Public Class _Constructor
            Inherits AccessPersisterTests

            ''' <summary>
            ''' Default constructor testing harness
            ''' </summary>
            <TestMethod>
            Public Sub DefaultConstructor()

                Dim obj As PFW.CSIST203.Project3.Persisters.Access.AccessPersister

                AssertDelegateSuccess(
                    Sub()
                        obj = New Project3.Persisters.Access.AccessPersister()
                    End Sub,
                    "Instantiation using the empty constructor should not throw an exception")

                Assert.IsTrue(String.IsNullOrWhiteSpace(obj.accessFile), "The access file variable should not be set when the parameterless constructor is used")
                Assert.IsTrue(obj.noDatabase, "The boolean value indicating no backing database is available should return true")

            End Sub

            ''' <summary>
            ''' Constructor passed with a file that does not exist
            ''' </summary>
            <TestMethod>
            Public Sub ConstructorWithMissingFile()

                Dim obj As PFW.CSIST203.Project3.Persisters.Access.AccessPersister
                Dim tmp = System.IO.Path.Combine(GetMethodSpecificWorkingDirectory(), System.Guid.NewGuid().ToString() & ".accdb")

                AssertDelegateFailure(
                    Sub()
                        obj = New Project3.Persisters.Access.AccessPersister(tmp)
                    End Sub,
                    GetType(System.IO.FileNotFoundException),
                    "An invalid file was supplied into the constructor and should have thrown a System.IO.FileNotFoundException")

                Assert.IsNull(obj, "The variable should still be null because it could not be instantiated")

            End Sub

            ''' <summary>
            ''' Constructor passed a path to an access database file that does exist
            ''' </summary>
            <TestMethod>
            Public Sub ConstructorWithValidFile()

                Dim obj As PFW.CSIST203.Project3.Persisters.Access.AccessPersister
                Dim directory = GetMethodSpecificWorkingDirectory()
                Dim tmpAccessDatabaseFile = System.IO.Path.Combine(directory, "sample-data.accdb")
                CopyEmbeddedResourceBaseToDirectory("PFW.CSIST203.Project3.Tests.Resources.Data", directory)
                Assert.IsTrue(System.IO.File.Exists(tmpAccessDatabaseFile), "The sample data access database file was not found")

                AssertDelegateSuccess(
                    Sub()
                        obj = New Project3.Persisters.Access.AccessPersister(tmpAccessDatabaseFile)
                    End Sub,
                    "Instantiation pointing to a file should not immediately throw an exception")

                Assert.AreEqual(tmpAccessDatabaseFile, obj.accessFile, "The local variable should be set to the supplied file in the constructor")
                Assert.IsFalse(obj.noDatabase, "A database file was supplied, so the boolean value should be set to false")

            End Sub

        End Class

        ''' <summary>
        ''' CountRows() method testing harness
        ''' </summary>
        <TestClass>
        Public Class CountRowsMethod
            Inherits AccessPersisterTests

            ''' <summary>
            ''' The parameterless constructor should return zero when this method is called
            ''' </summary>
            <TestMethod>
            Public Sub DefaultConstructedObjectReturnsZero()
                Dim obj = CreatePersister()
                Dim result = 0

                AssertDelegateSuccess(
                    Sub()
                        result = obj.CountRows()
                    End Sub,
                    "A persister created with the default constructor should not throw an exception when the CountRows() method is called")

                Assert.AreEqual(0, result, "A persister created with the default constructor should have no data")
            End Sub

            ''' <summary>
            ''' The method should return the current count of entries in the access database
            ''' </summary>
            <TestMethod>
            Public Sub RowCountIsCorrect()

                Dim obj = CreatePersister(GetMethodSpecificWorkingDirectory())
                Dim result = 0

                AssertDelegateSuccess(
                    Sub()
                        result = obj.CountRows()
                    End Sub,
                    "A call to the CountRows() method should not throw an exception when a valid database file was specified")

                Assert.AreEqual(9, result, "The sample data did not return the expected count")

            End Sub

            ''' <summary>
            ''' Calling this method after the Dispose() call should throw an exception
            ''' </summary>
            <TestMethod>
            Public Sub ThrowsObjectDisposedAfterDisposeCalled()

                Dim obj = CreatePersister()
                AssertDelegateSuccess(
                    Sub()
                        obj.Dispose()
                    End Sub,
                    "The Dispose() method should not throw an exception")

                AssertDelegateFailure(
                    Sub()
                        obj.CountRows()
                    End Sub,
                    GetType(System.ObjectDisposedException),
                    "A disposed object should throw an exception when the method is called")

            End Sub

        End Class

        ''' <summary>
        ''' Dispose() method testing harness
        ''' </summary>
        <TestClass>
        Public Class DisposeMethod
            Inherits AccessPersisterTests

            ''' <summary>
            ''' Calling this method should set the IsDisposed variable to True
            ''' </summary>
            <TestMethod>
            Public Sub DisposeSetsIsDisposedVariable()

                Dim obj = CreatePersister()
                AssertDelegateSuccess(
                    Sub()
                        obj.Dispose()
                    End Sub,
                    "The Dispose() method should not throw an exception")

                Assert.IsTrue(obj.IsDisposed, "Once Dispose() is called, the local IsDisposed variable should be true")

            End Sub

            ''' <summary>
            ''' Calling the Dispose() method multiple times should not throw an exception
            ''' </summary>
            <TestMethod>
            Public Sub MultipleDisposeCallsDoNotThrowException()

                Dim obj = CreatePersister()
                AssertDelegateSuccess(
                    Sub()
                        obj.Dispose()
                    End Sub,
                    "The Dispose() method should not throw an exception")
                Assert.IsTrue(obj.IsDisposed, "Once Dispose() is called, the local IsDisposed variable should be true")

                AssertDelegateSuccess(
                    Sub()
                        obj.Dispose()
                    End Sub,
                    "The Dispose() method should not throw an exception")
                Assert.IsTrue(obj.IsDisposed, "Once Dispose() is called, the local IsDisposed variable should be true")

                AssertDelegateSuccess(
                    Sub()
                        obj.Dispose()
                    End Sub,
                    "The Dispose() method should not throw an exception")
                Assert.IsTrue(obj.IsDisposed, "Once Dispose() is called, the local IsDisposed variable should be true")

            End Sub

        End Class

        ''' <summary>
        ''' GetRow() testing harness
        ''' </summary>
        <TestClass>
        Public Class GetRowMethod
            Inherits AccessPersisterTests

            ''' <summary>
            ''' Ensure the DataRow returned from the persister contains the correct data
            ''' </summary>
            <TestMethod>
            Public Sub ReadsDatabaseFileCorrectly()
                Dim obj = CreatePersister(GetMethodSpecificWorkingDirectory())
                Dim row As DataRow

                AssertDelegateSuccess(
                    Sub()
                        row = obj.GetRow(4)
                    End Sub,
                    "The GetRow() method should not have thrown an exception")

                Assert.AreEqual(4, row("ID"), "The database value in the sample file did not match the expected result")
                Assert.AreEqual("Northwind Traders", row("Company"), "The database value in the sample file did not match the expected result")
                Assert.AreEqual("Sergienko", row("Last Name"), "The database value in the sample file did not match the expected result")
                Assert.AreEqual("Mariya", row("First Name"), "The database value in the sample file did not match the expected result")
                Assert.AreEqual("mariya@northwindtraders.com", row("E-mail Address"), "The database value in the sample file did not match the expected result")
                Assert.AreEqual("Sales Representative", row("Job Title"), "The database value in the sample file did not match the expected result")
                Assert.AreEqual("(123)555-0104", row("Business Phone"), "The database value in the sample file did not match the expected result")


            End Sub

            ''' <summary>
            ''' Requesting data for an ID that does not exist should return a null (Nothing) DataRow
            ''' </summary>
            <TestMethod>
            Public Sub ReturnsNullIfNoData()
                Dim obj = CreatePersister()
                Dim row As DataRow

                AssertDelegateSuccess(
                    Sub()
                        row = obj.GetRow(123)
                    End Sub,
                    "A parameterless constructed persister should not throw an exception when calling GetRow()")

                Assert.IsNull(row, "Row object should have been null")

            End Sub

            ''' <summary>
            ''' If the Access Database is modified by another application the persister should pick up that change when it reloads the data
            ''' </summary>
            <TestMethod>
            Public Sub PicksUpExternalDatabaseChanges()
                Dim obj = CreatePersister(GetMethodSpecificWorkingDirectory())
                Dim row As DataRow

                AssertDelegateSuccess(
                    Sub()
                        row = obj.GetRow(2)
                    End Sub,
                    "Retrieval of a specific record in the access database should not throw an exception")

                ' ensure the state of the database
                Assert.AreEqual(2, row("ID"), "The database value in the sample file did not match the expected result")
                Assert.AreEqual("Northwind Traders", row("Company"), "The database value in the sample file did not match the expected result")
                Assert.AreEqual("Cencini", row("Last Name"), "The database value in the sample file did not match the expected result")
                Assert.AreEqual("Andrew", row("First Name"), "The database value in the sample file did not match the expected result")
                Assert.AreEqual("andrew@northwindtraders.com", row("E-mail Address"), "The database value in the sample file did not match the expected result")
                Assert.AreEqual("Vice President, Sales", row("Job Title"), "The database value in the sample file did not match the expected result")
                Assert.AreEqual("(123)555-0102", row("Business Phone"), "The database value in the sample file did not match the expected result")

                ' change a few values in the DataRow before persisting these changes back to the database
                row("Company") = "Acme Products, LLC"
                row("Business Phone") = "(260)123-4567"

                ' perform a manual modification to the record above that does not utilize the UpdateRow() method
                Using connection As New System.Data.OleDb.OleDbConnection(Util.GetAccessConnectionString(obj.accessFile))
                    connection.Open()
                    Using cmd = connection.CreateCommand()
                        cmd.CommandText = "UPDATE [tblEmployees] SET [Company] = @Company,[Last Name] = @LastName,[First Name] = @FirstName,[E-mail Address] = @EmailAddress,[Job Title] = @JobTitle,[Business Phone] = @BusinessPhone WHERE [ID] = @ID"

                        ' company parameter
                        Dim par = cmd.CreateParameter()
                        par.ParameterName = "@Company"
                        par.DbType = DbType.String
                        par.Value = row("Company")
                        cmd.Parameters.Add(par)

                        'last name
                        par = cmd.CreateParameter()
                        par.ParameterName = "@LastName"
                        par.DbType = DbType.String
                        par.Value = row("Last Name")
                        cmd.Parameters.Add(par)

                        'first name
                        par = cmd.CreateParameter()
                        par.ParameterName = "@FirstName"
                        par.DbType = DbType.String
                        par.Value = row("First Name")
                        cmd.Parameters.Add(par)

                        'e-mail address
                        par = cmd.CreateParameter()
                        par.ParameterName = "@EmailAddress"
                        par.DbType = DbType.String
                        par.Value = row("E-mail Address")
                        cmd.Parameters.Add(par)

                        'job title
                        par = cmd.CreateParameter()
                        par.ParameterName = "@JobTitle"
                        par.DbType = DbType.String
                        par.Value = row("Job Title")
                        cmd.Parameters.Add(par)

                        'business phone
                        par = cmd.CreateParameter()
                        par.ParameterName = "@BusinessPhone"
                        par.DbType = DbType.String
                        par.Value = row("Business Phone")
                        cmd.Parameters.Add(par)

                        par = cmd.CreateParameter()
                        par.ParameterName = "@ID"
                        par.DbType = DbType.Int32
                        par.Value = row("ID")
                        cmd.Parameters.Add(par)

                        Dim results = cmd.ExecuteNonQuery()

                        Assert.AreEqual(1, results, "The number of expected results from the database update was not returned")

                    End Using
                End Using

                ' Read the modified row using the persister
                AssertDelegateSuccess(
                    Sub()
                        row = obj.GetRow(2)
                    End Sub,
                    "Retrieval of a specific record in the access database should not throw an exception")

                Assert.AreEqual("Acme Products, LLC", row("Company"), "The database value in the sample file did not match the expected result after the external modification")
                Assert.AreEqual("(260)123-4567", row("Business Phone"), "The database value in the sample file did not match the expected result after the external modification")

            End Sub

            ''' <summary>
            ''' Calling this method should throw an ObjectDisposedException if the Dispose() method has been previously called
            ''' </summary>
            <TestMethod>
            Public Sub ThrowsObjectDisposedAfterDisposeCalled()

                Dim obj = CreatePersister()
                AssertDelegateSuccess(
                    Sub()
                        obj.Dispose()
                    End Sub,
                    "The Dispose() method should not throw an exception")

                AssertDelegateFailure(
                    Sub()
                        obj.GetRow(0)
                    End Sub,
                    GetType(System.ObjectDisposedException),
                    "A disposed object should throw an exception when the method is called")

            End Sub

        End Class

        ''' <summary>
        ''' GetData() method testing harness
        ''' </summary>
        <TestClass>
        Public Class GetDataMethod
            Inherits AccessPersisterTests

            ''' <summary>
            ''' The parameterless constructor should cause this method to return an empty data table
            ''' </summary>
            <TestMethod>
            Public Sub ParameterlessConstructorReturnsEmptyDataTable()
                Dim obj = CreatePersister()
                Dim data As System.Data.DataTable

                AssertDelegateSuccess(
                    Sub()
                        data = obj.GetData()
                    End Sub,
                    "Requesting all data contained in the underlying database file should not throw an exception")

                Assert.IsNotNull(data, "The DataTable returned by the method should be non-null")
                Assert.AreEqual(0, data.Rows.Count, "The number of data rows contained in a parameterless constructed persister should be zero")
            End Sub

            ''' <summary>
            ''' A call to this method should return a DataTable that contains all data from the Access Database table
            ''' </summary>
            <TestMethod>
            Public Sub AccessDatabaseFileReturnsAllResults()
                Dim obj = CreatePersister(GetMethodSpecificWorkingDirectory())
                Dim data As System.Data.DataTable

                AssertDelegateSuccess(
                    Sub()
                        data = obj.GetData()
                    End Sub,
                    "Requesting all data contained in the underlying database file should not throw an exception")

                Assert.IsNotNull(data, "The DataTable returned by the method should be non-null")

                ' perform a raw check of the testing database to ensure the persister isn't faking the results
                ' NOTE: This section may also fail if the persister does not Dispose() of any open OleDbConnection objects properly
                ' with a Using statement (as below) or by simply calling Dispose() on the object once finished
                Dim rawCount = 0
                Using connection As New System.Data.OleDb.OleDbConnection(Util.GetAccessConnectionString(obj.accessFile))
                    connection.Open()
                    Using cmd = connection.CreateCommand()
                        cmd.CommandText = "SELECT COUNT(*) FROM [tblEmployees]"
                        rawCount = Integer.Parse(cmd.ExecuteScalar().ToString().Trim())
                    End Using
                End Using

                Assert.AreEqual(rawCount, data.Rows.Count, "The persister and raw OleDb command check did not return identical results")

            End Sub

            ''' <summary>
            ''' Calling this method should throw an ObjectDisposedException if the Dispose() method has been previously called
            ''' </summary>
            <TestMethod>
            Public Sub ThrowsObjectDisposedAfterDisposeCalled()

                Dim obj = CreatePersister()
                AssertDelegateSuccess(
                    Sub()
                        obj.Dispose()
                    End Sub,
                    "The Dispose() method should not throw an exception")

                AssertDelegateFailure(
                    Sub()
                        obj.GetData()
                    End Sub,
                    GetType(System.ObjectDisposedException),
                    "A call into this method should throw a System.ObjectDisposedException")

            End Sub

        End Class

        ''' <summary>
        ''' UpdateRow() method testing harness
        ''' </summary>
        <TestClass>
        Public Class StoreRowMethod
            Inherits AccessPersisterTests

            ''' <summary>
            ''' Retrieval of an existing record, changing it and then updating it should return that value during subsequent loads
            ''' </summary>
            <TestMethod>
            Public Sub UpdateExistingRecordWorksWithPersister()
                Dim obj = CreatePersister(GetMethodSpecificWorkingDirectory())
                Dim row As System.Data.DataRow

                AssertDelegateSuccess(
                    Sub()
                        row = obj.GetRow(6)
                    End Sub,
                    "Retrieval of a specific record in the access database should not throw an exception")

                Assert.AreEqual("Sales Representative", row("Job Title"), "The expected value was not read from the sample access database")

                ' change their job title
                row("Job Title") = "Sales Manager"
                row("Business Phone") = "(123)555-5119"

                ' send the row to the access database to be updated
                AssertDelegateSuccess(
                    Sub()
                        obj.StoreRow(row)
                    End Sub,
                    "Updating a row should not throw an exception")

                ' ask the persister for the row again and ensure it matches the expected value
                AssertDelegateSuccess(
                    Sub()
                        row = obj.GetRow(6)
                    End Sub,
                    "Retrieval of a specific record in the access database should not throw an exception")

                Assert.AreEqual("Sales Manager", row("Job Title"), "The expected value was not read from the sample access database")
                Assert.AreEqual("(123)555-5119", row("Business Phone"), "The expected value was not read from the sample access database")

            End Sub

            ''' <summary>
            ''' Once the persister has updated a row, connect to the database manually and verify the write occurred as expected
            ''' </summary>
            <TestMethod>
            Public Sub UpdateExistingRecordWorksWithOleDb()

                Dim obj = CreatePersister(GetMethodSpecificWorkingDirectory())
                Dim row As System.Data.DataRow

                AssertDelegateSuccess(
                    Sub()
                        row = obj.GetRow(6)
                    End Sub,
                    "Retrieval of a specific record in the access database should not throw an exception")

                Assert.AreEqual("Sales Representative", row("Job Title"), "The expected value was not read from the sample access database")

                ' change their job title
                row("Job Title") = "Sales Supervisor"
                row("Business Phone") = "(123)555-6119"

                ' send the row to the access database to be updated
                AssertDelegateSuccess(
                    Sub()
                        obj.StoreRow(row)
                    End Sub,
                    "Updating a row should not throw an exception")

                ' retrieve the row in question using a raw OleDbConnection
                Dim dt As New System.Data.DataTable()
                Using connection As New System.Data.OleDb.OleDbConnection(Util.GetAccessConnectionString(obj.accessFile))
                    connection.Open()
                    Using cmd = connection.CreateCommand()
                        cmd.CommandText = "SELECT * FROM [tblEmployees] WHERE [ID] = @ID"

                        ' create the single parameter that selects the specific row we are interested in
                        Dim par = cmd.CreateParameter()
                        par.ParameterName = "@ID"
                        par.DbType = DbType.Int32
                        par.Value = 6
                        cmd.Parameters.Add(par)

                        Using dr = cmd.ExecuteReader()
                            dt.Load(dr)
                        End Using

                    End Using
                End Using

                Assert.AreEqual(1, dt.Rows.Count, "The sample data should only have returned a single result")
                Assert.AreEqual(row.Table.Columns.Count, dt.Columns.Count, "The column count from the two different retrieval methods did not match exactly")

                ' iterate over all of the columns owned by each of the data sources
                Dim OleDbMethodRow = dt.Rows(0)
                For Each column As System.Data.DataColumn In row.Table.Columns
                    Assert.AreEqual(row(column.ColumnName), OleDbMethodRow(column.ColumnName), "The column value updated did not match the value read from the access database for column: " & column.ColumnName)
                Next

            End Sub

            ''' <summary>
            ''' If an invalid ID is specified during the update operation an exception should be thrown indicating it failed
            ''' </summary>
            <TestMethod>
            Public Sub UpdateMissingRecordThrowsException()
                Dim obj = CreatePersister(GetMethodSpecificWorkingDirectory())
                Dim row As System.Data.DataRow

                AssertDelegateSuccess(
                    Sub()
                        row = obj.GetRow(6)
                    End Sub,
                    "Retrieval of a specific record in the access database should not throw an exception")

                ' change the ID (unique identifier) to an invalid value
                row("ID") = -25

                ' Try to update the value
                AssertDelegateFailure(
                    Sub()
                        obj.StoreRow(row)
                    End Sub,
                    GetType(System.ArgumentException),
                    "Updating an invalid row should throw a System.ArgumentException")

            End Sub

            ''' <summary>
            ''' Calling this method should throw an ObjectDisposedException if the Dispose() method has been previously called
            ''' </summary>
            <TestMethod>
            Public Sub ThrowsObjectDisposedAfterDisposeCalled()
                Dim obj = CreatePersister(GetMethodSpecificWorkingDirectory())
                Dim row As System.Data.DataRow

                AssertDelegateSuccess(
                    Sub()
                        row = obj.GetRow(7)
                    End Sub,
                    "Retrieval of a specific record in the access database should not throw an exception")

                AssertDelegateSuccess(
                    Sub()
                        obj.Dispose()
                    End Sub,
                    "The Dispose() method should not throw an exception")

                AssertDelegateFailure(
                    Sub()
                        obj.StoreRow(row)
                    End Sub,
                    GetType(System.ObjectDisposedException),
                    "A call into this method should throw a System.ObjectDisposedException")
            End Sub

            ''' <summary>
            ''' Verify that the persister can store new rows correctly in the access database
            ''' </summary>
            <TestMethod>
            Public Sub StoreNewRecordWorksCorrectly()

                Dim obj = CreatePersister(GetMethodSpecificWorkingDirectory())
                Dim row As System.Data.DataRow

                ' query the database and determine the number of rows present in the database
                Dim currentcount = obj.CountRows()

                AssertDelegateSuccess(
                    Sub()
                        row = obj.CreateRow("tblEmployees")
                    End Sub,
                    "Creating a new row object should not throw an exception")
                Assert.IsNotNull(row, "The row object cannot be null")

                ' make sure not new entry was created by simply calling the CreateRow() method
                Assert.AreEqual(currentcount, obj.CountRows(), "Simply calling the CreateRow() method should NOT have created an entry in the database")

                ' assign the values for the new row
                'Company
                row("Company") = "Acme, Inc"
                'Last Name
                row("Last Name") = "Smith"
                'First Name
                row("First Name") = "Joe"
                'E-mail Address
                row("E-mail Address") = "jsmith@acme.org"
                'Job Title
                row("Job Title") = "Senior Marketing Specialist"
                'Business Phone
                row("Business Phone") = "(742)555-5555"

                AssertDelegateSuccess(
                    Sub()
                        obj.StoreRow(row)
                    End Sub,
                    "Storing a new entry into the database should not fail")

                Assert.AreEqual(currentcount + 1, obj.CountRows(), "A new row was not created after the successful save operation")

                ' make sure the 'SELECT @@IDENTITY' initialized the ID column
                Assert.IsNotNull(row("ID"), "The ID value was not retrieved propertly from the database after the store operation")
                Assert.AreNotEqual(row("ID"), 0, "The ID must be a positive integer if set propertly by the SELECT @@IDENTITY sql query")
                Assert.IsFalse(row.IsNull("ID"), "The ID value was not retrieved propertly from the database after the store operation")

                Dim dt As New System.Data.DataTable()
                Using connection As New System.Data.OleDb.OleDbConnection(Util.GetAccessConnectionString(obj.accessFile))
                    connection.Open()
                    Using cmd = connection.CreateCommand()
                        cmd.CommandText = "SELECT * FROM [tblEmployees] WHERE [ID] = @ID"

                        ' create the single parameter that selects the specific row we are interested in
                        Dim par = cmd.CreateParameter()
                        par.ParameterName = "@ID"
                        par.DbType = DbType.Int32
                        par.Value = row("ID")
                        cmd.Parameters.Add(par)

                        Using dr = cmd.ExecuteReader()
                            dt.Load(dr)
                        End Using

                    End Using
                End Using

                Assert.AreEqual(1, dt.Rows.Count, "The raw SQL query should have returned exactly one row")

                ' make sure all values that were saved into the database match excactly with the values that were found using raw SQL queries
                Dim newRow = CType(dt.Rows(0), DataRow)
                For Each column As DataColumn In newRow.Table.Columns
                    Assert.AreEqual(row(column.ColumnName), newRow(column.ColumnName),
                        String.Format("The column ('{0}') value from the database ('{1}') did not match the value that was supplied during the store operation ('{2}')", column.ColumnName, newRow(column.ColumnName), row(column.ColumnName)))
                Next

            End Sub

        End Class

        <TestClass>
        Public Class CreateRowMethod
            Inherits AccessPersisterTests

            <TestMethod>
            Public Sub SupportsEmployeeTable()
                Dim obj = CreatePersister(GetMethodSpecificWorkingDirectory())
                Dim row As System.Data.DataRow

                AssertDelegateSuccess(
                    Sub()
                        row = obj.CreateRow("tblEmployees")
                    End Sub,
                    "Creation of a row for the tblEmployees table should always work")

                Assert.IsNotNull(row, "The row object returned should be non-null")

                AssertDelegateFailure(
                    Sub()
                        row = obj.CreateRow(System.Guid.NewGuid().ToString())
                    End Sub,
                    GetType(ArgumentException),
                    "The persister should throw an exception when an invalid table is specified")

            End Sub

            <TestMethod>
            Public Sub NewRowHasCorrectColumns()
                Dim obj = CreatePersister(GetMethodSpecificWorkingDirectory())
                Dim row As System.Data.DataRow

                AssertDelegateSuccess(
                    Sub()
                        row = obj.CreateRow("tblEmployees")
                    End Sub,
                    "Creation of a row for the tblEmployees table should always work")

                Dim columns = row.Table.Columns.Cast(Of DataColumn)
                Dim expectedColumnNames As New List(Of String) From {"ID", "Company", "Last Name", "First Name", "E-mail Address", "Job Title", "Business Phone"}
                Dim expectedColumnTypes As New List(Of Type) From {GetType(Integer), GetType(String), GetType(String), GetType(String), GetType(String), GetType(String), GetType(String)}

                For i As Integer = 0 To expectedColumnNames.Count - 1

                    Dim columnName = expectedColumnNames(i)
                    Dim column = columns.FirstOrDefault(Function(c) String.Equals(c.ColumnName, columnName, StringComparison.OrdinalIgnoreCase))
                    Assert.IsNotNull(column, String.Format("Column '{0}' missing from the newly created data row", columnName))
                    Assert.AreEqual(expectedColumnTypes(i), column.DataType, "The data type of the column ({0}) did not match the expected type ({1})", expectedColumnTypes(i), column.DataType)

                Next

            End Sub

            <TestMethod>
            Public Sub ThrowsObjectDisposedAfterDisposeCalled()
                Dim obj = CreatePersister(GetMethodSpecificWorkingDirectory())
                Dim row As System.Data.DataRow

                AssertDelegateSuccess(
                    Sub()
                        row = obj.CreateRow("tblEmployees")
                    End Sub,
                    "Creating a new row object should not throw an exception")

                AssertDelegateSuccess(
                    Sub()
                        obj.Dispose()
                    End Sub,
                    "The Dispose() method should not throw an exception")

                AssertDelegateFailure(
                    Sub()
                        row = obj.CreateRow("tblEmployees")
                    End Sub,
                    GetType(System.ObjectDisposedException),
                    "A call into this method should throw a System.ObjectDisposedException")
            End Sub

        End Class


    End Class

End Namespace

