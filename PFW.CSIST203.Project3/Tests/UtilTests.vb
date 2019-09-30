Imports Microsoft.VisualStudio.TestTools.UnitTesting

Namespace PFW.CSIST203.Project3.Tests

    ''' <summary>
    ''' Testing harness for the Util class
    ''' </summary>
    Public MustInherit Class UtilTests
        Inherits TestBase

        ''' <summary>
        ''' GetExcelConnectionString() method testing harness
        ''' </summary>
        <TestClass>
        Public Class GetExcelConnectionStringMethod
            Inherits UtilTests

            ''' <summary>
            ''' Verify a correct connection string for 2007 on onward file format with header option specified as yes
            ''' </summary>
            <TestMethod>
            Public Sub Excel2007FilenameWithHeader()
                Dim connectionString = Util.GetExcelConnectionString("bogus.xlsx", True)
                Assert.IsTrue(connectionString.IndexOf("HDR=Yes", StringComparison.OrdinalIgnoreCase) >= 0, "Header option not specified in excel connection string")
            End Sub

            ''' <summary>
            ''' Verify a correct connection string for pre-2007 excel file format with header option specified as no
            ''' </summary>
            <TestMethod>
            Public Sub Excel2007FilenameWithoutHeader()
                Dim connectionString = Util.GetExcelConnectionString("bogus.xlsx", False)
                Assert.IsTrue(connectionString.IndexOf("HDR=No", StringComparison.OrdinalIgnoreCase) >= 0, "Header option not specified in excel connection string")
            End Sub

            ''' <summary>
            ''' Utility method that verifies the current machine can actually use the Microsoft Excel provider specified in the connection string
            ''' </summary>
            <TestMethod>
            Public Sub ExcelOleDb12ProviderIsRegistereOnLocalMachine()

                Dim directory = GetMethodSpecificWorkingDirectory()
                Dim tmpExcelFile = System.IO.Path.Combine(directory, "data005.xlsx")
                CopyEmbeddedResourceBaseToDirectory("PFW.CSIST203.Project3.Tests.Resources.Util.ExcelOleDb12ProviderIsRegistereOnLocalMachine", directory)
                Assert.IsTrue(System.IO.File.Exists(tmpExcelFile), "Unable to extract testing excel file from the embedded assembly")

                Using table = New System.Data.DataTable("Sheet1")
                    Try
                        Using connection As New System.Data.OleDb.OleDbConnection(Util.GetExcelConnectionString(tmpExcelFile, True))
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

                        logger.Error("Problem reading excel file: " + tmpExcelFile, ex)
                        Throw
                    End Try
                End Using

            End Sub

        End Class

        <TestClass>
        Public Class GetAccessConnectionString
            Inherits UtilTests

            ''' <summary>
            ''' Utility method that verifies the current machine can actually use the Microsoft Access provider specified in the connection string
            ''' </summary>
            <TestMethod>
            Public Sub AccessOleDb12ProviderIsRegistereOnLocalMachine()

                Dim directory = GetMethodSpecificWorkingDirectory()
                Dim tmpAccessDatabase = System.IO.Path.Combine(directory, "sample-data.accdb")
                CopyEmbeddedResourceBaseToDirectory("PFW.CSIST203.Project3.Tests.Resources.Data", directory)
                Assert.IsTrue(System.IO.File.Exists(tmpAccessDatabase), "Unable to extract testing access database file from the embedded assembly")

                Using table = New System.Data.DataTable("tblEmployees")
                    Try
                        Using connection As New System.Data.OleDb.OleDbConnection(Util.GetAccessConnectionString(tmpAccessDatabase))
                            Using cmd As System.Data.IDbCommand = connection.CreateCommand()
                                cmd.CommandText = "SELECT * FROM [tblEmployees]"
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

                        logger.Error("Problem reading access database file: " + tmpAccessDatabase, ex)
                        Throw
                    End Try
                End Using

            End Sub

        End Class

    End Class

End Namespace


