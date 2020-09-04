Imports CommandLine
Public Class Options
    <CommandLine.Option("s", "sla", Default:=24, HelpText:="Input Service Level Agreement")>
    Public Property SLA As Integer

    <CommandLine.Option("l", "location", Required:=True, HelpText:="Input Location")>
    Public Property Location As String

    <CommandLine.Option("d", "datetime", Required:=True, HelpText:="Input Date and Time")>
    Public Property DateTime As String
End Class
