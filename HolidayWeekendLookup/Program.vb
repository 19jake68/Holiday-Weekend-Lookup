Imports CommandLine

Module Program
    Dim ExcelModel = New ExcelModel

    Sub Main(ByVal args As String())
        EnsureInitialized()
        Dim Options = New Options

        Parser.Default.ParseArguments(Of Options)(args) _
        .WithParsed(Sub(opts As Options)
                        Console.WriteLine(_EndedOnHolidayOrWeekend(opts))
                    End Sub) _
        .WithNotParsed(Function(errs As IEnumerable(Of [Error])) 1)
    End Sub

    Private Function _EndedOnHolidayOrWeekend(ByVal opts As Options) As Boolean
        Dim SLA As Integer = opts.SLA
        Dim Location As String = opts.Location
        Dim DateTime As DateTime = opts.DateTime
        Dim DateEndOfSLA As DateTime = DateTime.AddHours(SLA)

        ' Check if date when the SLA will run out is a holiday or a weekend.
        Return _IsWeekend(DateEndOfSLA) Or _IsHoliday(DateEndOfSLA, Location)
    End Function

    Private Function _IsWeekend(ByVal dt As DateTime) As Boolean
        Return dt.DayOfWeek = DayOfWeek.Saturday Or dt.DayOfWeek = DayOfWeek.Sunday
    End Function

    Private Function _IsHoliday(ByVal Dt As DateTime, ByVal Loc As String) As Boolean
        ExcelModel.Open
        Return ExcelModel.CheckIfHoliday(Dt.ToShortDateString, Loc)
    End Function
End Module