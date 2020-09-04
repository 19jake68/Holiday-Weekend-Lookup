Imports CommandLine

Module Program
    Sub Main(ByVal args As String())
        Dim Options = New Options

        Dim unused = CommandLine.Parser.Default.ParseArguments(Of Options)(args) _
        .WithParsed(Function(opts As Options) _EndedOnHolidayOrWeekend(opts)) _
        .WithNotParsed(Function(errs As IEnumerable(Of [Error])) 1)
    End Sub

    Private Function _EndedOnHolidayOrWeekend(ByVal opts As Options)
        Dim SLA = opts.SLA
        Dim Location = opts.Location
        Dim DateTime = opts.DateTime

        Console.WriteLine(SLA)
        Console.WriteLine(Location)
        Console.WriteLine(DateTime)

        Return 0
    End Function
End Module
