Imports System.Configuration
Imports System.Data
Imports System.IO
Imports System.Reflection
Imports System.Text
Imports ExcelDataReader

Public Class ExcelModel
    Private DTCollection As DataTableCollection

    Public Sub Connect()
        Try
            ' Get file full location and read it
            Dim AssemblyName As String = Assembly.GetExecutingAssembly.GetModules()(0).FullyQualifiedName
            Dim FilePath As String = Path.GetDirectoryName(AssemblyName) & ConfigurationManager.AppSettings.Get("ExcelFile")
            Dim FStream As FileStream = File.Open(FilePath, FileMode.Open, FileAccess.Read)

            ' Create Excel reader from stream
            Dim ExcelDataReader As IExcelDataReader
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance)
            If Path.GetExtension(FilePath).ToUpper() = ".XLS" Then
                ExcelDataReader = ExcelReaderFactory.CreateReader(FStream)
            Else
                ExcelDataReader = ExcelReaderFactory.CreateOpenXmlReader(FStream)
            End If

            ' Configure Excel to DataSet
            Dim Result As DataSet = ExcelDataReader.AsDataSet(New ExcelDataSetConfiguration() With {
                .ConfigureDataTable = Function(__) New ExcelDataTableConfiguration() With {
                .UseHeaderRow = True}})
            DTCollection = Result.Tables()
        Catch ex As Exception
            Console.WriteLine("Exception: " & ex.ToString())
            Environment.Exit(-1)
        End Try
    End Sub

    Public Function CheckIfHoliday(ByVal Dt As Date, ByVal Loc As String)
        Dim Sheetname = Dt.Year.ToString
        For Each Table As DataTable In DTCollection
            If (Table.TableName = Sheetname) Then
                Dim DRow As DataRow() = Table.Select("Date = '" & Dt & "' AND Country = '" & Loc & "'")
                Return If(DRow.Length > 0, True, False)
            End If
        Next

        Return False
    End Function
End Class
