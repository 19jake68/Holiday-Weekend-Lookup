Imports System.Configuration
Imports System.Data
Imports System.IO
Imports System.Reflection
Imports System.Text
Imports ExcelDataReader

Public Class ExcelConnection
    ' Constructor
    Public Sub New()
        Try
            ' Get file full location and read it
            Dim assemblyName As String = Assembly.GetExecutingAssembly.GetModules()(0).FullyQualifiedName
            Dim filePath As String = Path.GetDirectoryName(assemblyName) & ConfigurationManager.AppSettings.Get("ExcelFile")
            Dim stream As FileStream = File.Open(filePath, FileMode.Open, FileAccess.Read)

            ' Create Excel reader from stream
            Dim excelReader As IExcelDataReader
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance)
            If Path.GetExtension(filePath).ToUpper() = ".XLS" Then
                excelReader = ExcelReaderFactory.CreateReader(stream)
            Else
                excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream)
            End If

            ' Configure Excel to DataSet
            Dim result As DataSet = excelReader.AsDataSet(New ExcelDataSetConfiguration() With {
                .ConfigureDataTable = Function(__) New ExcelDataTableConfiguration() With {
                .UseHeaderRow = True}})
            Dim dataTable As DataTableCollection = result.Tables

            ' Return datatable
            For Each table As DataTable In dataTable
                Console.WriteLine(table.TableName)
            Next
        Catch ex As Exception
            Console.WriteLine("Exception: " & ex.ToString())
            Environment.Exit(-1)
        End Try
    End Sub


End Class
